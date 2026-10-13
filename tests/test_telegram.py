"""Offline threshold, date and safe Telegram integration tests."""

import io
import json
import os
import ssl
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import Mock, patch
from urllib.error import HTTPError, URLError

from main import audit_csv, main
from telegram_alerts import Alert, NoRedirects, TelegramError, api_call, credentials, messages, send_alerts
from telegram_setup import main as setup_main


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates/2026_PowerUsageTable_template.xlsx"
TOKEN = "123456789:" + "X" * 35  # Deliberately fake; all HTTP is mocked.


class TelegramTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.inputs = self.folder / "input"
        self.inputs.mkdir()
        self.output = self.folder / "output/report.xlsx"
        excel_patch = patch("main.send_excel")
        self.excel_sender = excel_patch.start()
        self.addCleanup(excel_patch.stop)

    def tearDown(self):
        self.temp.cleanup()

    def report(self, filename="export.csv", *, source="PDU_2A1", day="06.10.2026",
               average=0.02, maximum=0.35):
        pdu = 2 if "A" in source else 1
        path = self.inputs / filename
        path.write_text(f"{source};\nslot:/Services/ReportService/{source};\n{day};\n"
                        f"11A_PDU{pdu}_average;11A_PDU{pdu}_max\n{average};{maximum}\n",
                        encoding="utf-8-sig")
        return path

    def args(self, *extra):
        return ["--input-dir", str(self.inputs), "--template", str(TEMPLATE),
                "--output", str(self.output), *extra]

    def invoke(self, *extra):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            result = main(self.args(*extra))
        return result, out.getvalue(), err.getvalue()

    def test_only_max_is_checked_and_boundaries_are_inclusive(self):
        for maximum, count in ((0.0, 1), (0.099, 1), (0.1, 0), (4.0, 0), (4.001, 1)):
            with self.subTest(maximum=maximum):
                alerts = audit_csv(self.report(maximum=maximum))
                self.assertEqual(len(alerts), count)
                if alerts:
                    self.assertIn("Max:", alerts[0].description)
                    self.assertEqual(alerts[0].source, "PDU_2A1")
        self.assertEqual(audit_csv(self.report(average=0.001, maximum=3)), [])

    def test_original_inputs_generate_no_date_or_power_alert(self):
        for path in (ROOT / "tests/fixtures/original").glob("*.csv"):
            self.assertEqual(audit_csv(path), [])

    def test_calendar_dates_formats_and_leap_years(self):
        for day in ("31.02.2026", "00.10.2026", "29.02.2025", "2026-13-01",
                    "6.10.2026", "06/10/2026"):
            with self.subTest(day=day):
                alerts = audit_csv(self.report(day=day))
                self.assertEqual(len(alerts), 1)
                self.assertIn("Некорректная дата", alerts[0].description)
        for day in ("29.02.2024", "2026-10-06", "06.10.2026"):
            self.assertEqual(audit_csv(self.report(day=day)), [])

    def test_missing_date_and_non_numeric_max_are_reported(self):
        alerts = audit_csv(self.report(day="", maximum="abc"))
        self.assertEqual(len(alerts), 2)
        self.assertIn("дата", alerts[0].description)
        self.assertIn("Max", alerts[1].description)

    def test_decimal_comma_and_max_column_reordering(self):
        path = self.report()
        path.write_text("PDU_2A1;\n06.10.2026;\n11A_PDU2_max;11A_PDU2_average\n4,1;0,05\n",
                        encoding="cp1251")
        alerts = audit_csv(path)
        self.assertEqual(len(alerts), 1)
        self.assertIn("4.1 кВт", alerts[0].description)

    def test_all_blocks_in_file_are_audited_even_after_bad_date(self):
        path = self.report(day="31.02.2026", maximum=5)
        path.write_text(path.read_text(encoding="utf-8-sig") +
                        "PDU_3B8;\n31.04.2026;\nA7_PDU1_max;A7_PDU1_average\n0.09;0.01\n",
                        encoding="utf-8-sig")
        alerts = audit_csv(path)
        self.assertEqual(len(alerts), 4)
        self.assertEqual({a.source for a in alerts}, {"PDU_2A1", "PDU_3B8"})

    def test_date_reset_when_source_section_changes(self):
        path = self.report()
        path.write_text(path.read_text(encoding="utf-8-sig") +
                        "PDU_3B8;\nA7_PDU1_max;A7_PDU1_average\n1;0.1\n",
                        encoding="utf-8-sig")
        alerts = audit_csv(path)
        self.assertEqual(len(alerts), 1)
        self.assertEqual(alerts[0].source, "PDU_3B8")

    def test_errors_in_all_files_send_alerts_and_preserve_previous_excel(self):
        self.report("a.csv", day="31.02.2026", maximum=5)
        self.report("b.csv", source="PDU_3B8", day="31.04.2026", maximum=0.09)
        self.output.parent.mkdir()
        self.output.write_bytes(b"previous valid report")
        with patch.dict(os.environ, {"TELEGRAM_BOT_TOKEN": TOKEN, "TELEGRAM_CHAT_ID": "12345"}), \
                patch("main.send_alerts", return_value=2) as sender:
            status, _, error = self.invoke("--telegram")
        self.assertEqual(status, 1)
        self.assertIn("a.csv", error)
        self.assertIn("b.csv", error)
        self.assertEqual(len(sender.call_args.args[0]), 4)
        self.assertEqual(self.output.read_bytes(), b"previous valid report")
        self.assertIn("PDU_3B8:", (self.output.parent / "alerts.txt").read_text(encoding="utf-8"))

    def test_preview_and_check_make_no_network_requests_or_files(self):
        self.report(maximum=5)
        with patch("telegram_alerts.api_call") as api:
            status, output, _ = self.invoke("--check", "--alerts-preview")
        api.assert_not_called()
        self.assertEqual(status, 0)
        self.assertIn("PDU_2A1:", output)
        self.assertFalse(self.output.parent.exists())

    def test_without_explicit_flag_credentials_do_not_enable_sending(self):
        self.report(maximum=5)
        with patch.dict(os.environ, {"TELEGRAM_BOT_TOKEN": TOKEN, "TELEGRAM_CHAT_ID": "12345"}), \
                patch("main.send_alerts") as sender:
            status, _, _ = self.invoke()
        sender.assert_not_called()
        self.assertEqual(status, 0)
        self.assertTrue(self.output.exists())

    def test_delivery_failure_leaves_excel_and_local_alerts(self):
        self.report(maximum=5)
        with patch.dict(os.environ, {"TELEGRAM_BOT_TOKEN": TOKEN, "TELEGRAM_CHAT_ID": "12345"}), \
                patch("main.send_alerts", side_effect=TelegramError("Safe delivery failure")):
            status, _, error = self.invoke("--telegram")
        self.assertEqual(status, 2)
        self.assertNotIn(TOKEN, error)
        self.assertTrue(self.output.exists())
        self.assertIn("Max: 5", (self.output.parent / "alerts.txt").read_text(encoding="utf-8"))

    def test_bad_configuration_does_not_write_or_send(self):
        self.report(maximum=5)
        with patch.dict(os.environ, {"TELEGRAM_BOT_TOKEN": "", "TELEGRAM_CHAT_ID": ""}), \
                patch("main.send_alerts") as sender:
            status, _, _ = self.invoke("--telegram")
        self.assertEqual(status, 2)
        sender.assert_not_called()
        self.assertFalse(self.output.exists())

    def test_env_file_is_literal_and_cannot_override_destination_from_environment(self):
        config = self.folder / "telegram.env"
        config.write_text(f"TELEGRAM_BOT_TOKEN={TOKEN}\nTELEGRAM_CHAT_ID=-12345\n")
        config.chmod(0o600)
        with patch.dict(os.environ, {"TELEGRAM_CHAT_ID": "67890"}):
            self.assertEqual(credentials(config), (TOKEN, "-12345"))
        config.write_text(f"TELEGRAM_BOT_TOKEN={TOKEN}\nTELEGRAM_CHAT_ID=$(execute)\n")
        with self.assertRaises(TelegramError) as caught:
            credentials(config)
        self.assertNotIn(TOKEN, str(caught.exception))

    def test_long_reports_split_and_repeat_source_heading(self):
        alerts = [Alert("PDU_2A1", "a.csv", i, "Max " + "😀" * 300) for i in range(30)]
        texts = messages(alerts)
        self.assertGreater(len(texts), 1)
        for text in texts:
            self.assertTrue(text.startswith("PDU_2A1:\n"))
            self.assertLessEqual(len(text.encode("utf-16-le")) // 2, 4096)
        self.assertEqual(messages([alerts[0], alerts[0]]), messages([alerts[0]]))

    def test_https_plain_text_body_and_bounded_timeout(self):
        response = Mock()
        response.__enter__ = Mock(return_value=response)
        response.__exit__ = Mock(return_value=False)
        response.read.return_value = b'{"ok":true,"result":{"message_id":1}}'
        opener = Mock()
        opener.open.return_value = response
        with patch("telegram_alerts.build_opener", return_value=opener) as build:
            api_call(TOKEN, "sendMessage", {"chat_id": "12345", "text": "PDU_2A1:\n<text>"})
        req = opener.open.call_args.args[0]
        self.assertTrue(req.full_url.startswith("https://api.telegram.org/"))
        self.assertNotIn("parse_mode", json.loads(req.data))
        self.assertEqual(opener.open.call_args.kwargs["timeout"], 15)
        context = build.call_args.args[1]._context
        self.assertEqual(context.verify_mode, ssl.CERT_REQUIRED)
        self.assertTrue(context.check_hostname)
        self.assertIsInstance(build.call_args.args[0], NoRedirects)

    def test_http_network_and_invalid_responses_never_leak_token_or_retry(self):
        failures = [HTTPError(f"https://x/bot{TOKEN}", 401, TOKEN, {}, io.BytesIO()),
                    URLError(TOKEN)]
        for failure in failures:
            opener = Mock()
            opener.open.side_effect = failure
            with patch("telegram_alerts.build_opener", return_value=opener):
                with self.assertRaises(TelegramError) as caught:
                    api_call(TOKEN, "sendMessage", {})
            self.assertNotIn(TOKEN, str(caught.exception))
            self.assertEqual(opener.open.call_count, 1)
        self.assertIsNone(NoRedirects().redirect_request(None, None, 302, "", {}, "https://evil"))

    def test_partial_delivery_is_reported_without_credentials(self):
        alerts = [Alert("PDU_2A1", "a.csv", 1, "Max 5"),
                  Alert("PDU_2B1", "b.csv", 1, "Max 6")]
        with patch("telegram_alerts.api_call", side_effect=[{}, TelegramError("Failed")]), \
                patch("telegram_alerts.time.sleep"):
            with self.assertRaisesRegex(TelegramError, "Confirmed sent: 1"):
                send_alerts(alerts, TOKEN, "12345")

    def test_invalid_or_rejected_api_response_is_sanitized(self):
        for raw in (b"not json", json.dumps({"ok": False, "description": TOKEN}).encode(),
                    b'[]', b"x" * 1_000_001):
            response = Mock()
            response.__enter__ = Mock(return_value=response)
            response.__exit__ = Mock(return_value=False)
            response.read.return_value = raw
            opener = Mock()
            opener.open.return_value = response
            with patch("telegram_alerts.build_opener", return_value=opener):
                with self.assertRaises(TelegramError) as caught:
                    api_call(TOKEN, "sendMessage", {})
            self.assertNotIn(TOKEN, str(caught.exception))

    def test_no_warnings_means_no_send_request(self):
        self.report()
        with patch.dict(os.environ, {"TELEGRAM_BOT_TOKEN": TOKEN, "TELEGRAM_CHAT_ID": "12345"}), \
                patch("telegram_alerts.api_call") as api:
            status, output, _ = self.invoke("--telegram")
        self.assertEqual(status, 0)
        self.assertIn("Telegram messages sent: 0", output)
        api.assert_not_called()

    def test_group_sends_are_spaced_to_respect_group_limits(self):
        alerts = [Alert("PDU_2A1", "a.csv", 1, "Max 5"),
                  Alert("PDU_2B1", "b.csv", 1, "Max 6")]
        with patch("telegram_alerts.api_call", return_value={}), \
                patch("telegram_alerts.time.sleep") as sleep:
            self.assertEqual(send_alerts(alerts, TOKEN, "-12345"), 2)
        sleep.assert_called_once_with(3.1)

    def test_chat_setup_does_not_send_or_print_token(self):
        out = io.StringIO()
        with patch.dict(os.environ, {"TELEGRAM_BOT_TOKEN": TOKEN, "TELEGRAM_CHAT_ID": ""}), \
                patch("telegram_setup.api_call", return_value=[
                    {"message": {"chat": {"id": 12345, "type": "private"}}}]) as api, \
                redirect_stdout(out):
            self.assertEqual(setup_main([]), 0)
        self.assertIn("12345", out.getvalue())
        self.assertNotIn(TOKEN, out.getvalue())
        self.assertEqual(api.call_args.args[1], "getUpdates")


if __name__ == "__main__":
    unittest.main()
