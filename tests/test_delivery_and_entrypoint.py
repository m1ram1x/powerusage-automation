"""Date-based names, Excel delivery and private mounted configuration handling."""

import io
import json
import os
import tempfile
import unittest
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path
from unittest.mock import Mock, patch
from urllib.error import URLError

import container_entrypoint
from main import main
from telegram_alerts import TelegramError, send_excel


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates/2026_PowerUsageTable_template.xlsx"
TOKEN = "123456789:" + "X" * 35  # Fake token; no real network calls.


class DeliveryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.inputs = self.folder / "input"
        self.inputs.mkdir()

    def tearDown(self):
        self.temp.cleanup()

    def report(self, filename, day):
        (self.inputs / filename).write_text(
            f"PDU_2A1;\n{day};\n11A_PDU2_max;11A_PDU2_average\n0.35;0.24\n",
            encoding="utf-8-sig")

    def invoke(self, *extra):
        arguments = ["--input-dir", str(self.inputs), "--template", str(TEMPLATE), *extra]
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            status = main(arguments)
        return status, out.getvalue(), err.getvalue()

    def test_dates_drive_name_and_same_day_rerun_replaces_file(self):
        self.report("first.csv", "06.10.2026")
        old_cwd = Path.cwd()
        try:
            os.chdir(self.folder)
            status, _, _ = self.invoke()
            output = self.folder / "output/PowerUsage_2026-10-06.xlsx"
            self.assertEqual(status, 0)
            self.assertTrue(output.is_file())
            output.write_bytes(b"manually changed old output")
            self.assertEqual(self.invoke()[0], 0)
            self.assertTrue(output.read_bytes().startswith(b"PK"))
            self.report("second.csv", "07.10.2026")
            self.assertEqual(self.invoke()[0], 0)
            self.assertTrue((self.folder / "output/PowerUsage_2026-10-06_to_2026-10-07.xlsx").is_file())
        finally:
            os.chdir(old_cwd)

    def test_inline_output_option_is_not_overridden(self):
        self.report("first.csv", "06.10.2026")
        output = self.folder / "specific.xlsx"
        self.assertEqual(self.invoke(f"--output={output}")[0], 0)
        self.assertTrue(output.is_file())

    def test_excel_is_sent_without_warnings_and_not_after_data_failure(self):
        self.report("first.csv", "06.10.2026")
        output = self.folder / "report.xlsx"
        with patch.dict(os.environ, {"TELEGRAM_BOT_TOKEN": TOKEN, "TELEGRAM_CHAT_ID": "12345"}), \
                patch("main.send_alerts", return_value=0) as alerts, patch("main.send_excel") as excel:
            self.assertEqual(self.invoke("--output", str(output), "--telegram")[0], 0)
            alerts.assert_called_once_with([], TOKEN, "12345")
            excel.assert_called_once_with(output, TOKEN, "12345")
        self.report("first.csv", "31.02.2026")
        with patch.dict(os.environ, {"TELEGRAM_BOT_TOKEN": TOKEN, "TELEGRAM_CHAT_ID": "12345"}), \
                patch("main.send_alerts", return_value=1), patch("main.send_excel") as excel:
            self.assertEqual(self.invoke("--output", str(output), "--telegram")[0], 1)
            excel.assert_not_called()

    def test_document_request_contains_file_and_never_embeds_token_in_body(self):
        output = self.folder / "report.xlsx"
        output.write_bytes(b"PK-FAKE-EXCEL")
        response = Mock()
        response.__enter__ = Mock(return_value=response)
        response.__exit__ = Mock(return_value=False)
        response.read.return_value = b'{"ok":true,"result":{"message_id":1}}'
        opener = Mock()
        opener.open.return_value = response
        with patch("telegram_alerts.build_opener", return_value=opener), patch("telegram_alerts.time.sleep"):
            send_excel(output, TOKEN, "12345")
        request = opener.open.call_args.args[0]
        self.assertTrue(request.full_url.endswith("/sendDocument"))
        self.assertIn(b'name="document"', request.data)
        self.assertIn(b"PK-FAKE-EXCEL", request.data)
        self.assertNotIn(TOKEN.encode(), request.data)

    def test_document_failure_does_not_expose_token(self):
        output = self.folder / "report.xlsx"
        output.write_bytes(b"PK-FAKE-EXCEL")
        opener = Mock()
        opener.open.side_effect = URLError(TOKEN)
        with patch("telegram_alerts.build_opener", return_value=opener), patch("telegram_alerts.time.sleep"):
            with self.assertRaises(TelegramError) as caught:
                send_excel(output, TOKEN, "12345")
        self.assertNotIn(TOKEN, str(caught.exception))

    def test_entrypoint_copies_only_to_owner_private_temp_and_cleans_it(self):
        source = self.folder / "fake.env"
        source.write_text("TELEGRAM_BOT_TOKEN=FAKE\nTELEGRAM_CHAT_ID=12345\n")
        seen = []

        def inspect(arguments):
            copied = Path(arguments[arguments.index("--telegram-env") + 1])
            self.assertNotEqual(copied, source)
            self.assertEqual(copied.read_bytes(), source.read_bytes())
            if os.name != "nt":
                self.assertEqual(copied.stat().st_mode & 0o777, 0o600)
            seen.append(copied)
            return 0

        with patch("container_entrypoint.process_batch", side_effect=inspect):
            self.assertEqual(container_entrypoint.main(["--telegram", "--telegram-env", str(source)]), 0)
        self.assertFalse(seen[0].exists())
        self.assertTrue(source.exists())

    def test_setup_chat_routes_to_helper_and_cleans_its_private_copy(self):
        source = self.folder / "fake.env"
        source.write_text("TELEGRAM_BOT_TOKEN=FAKE\nTELEGRAM_CHAT_ID=\n")
        with patch("container_entrypoint.find_chat_ids", return_value=0) as setup, \
                patch("container_entrypoint.process_batch") as process:
            self.assertEqual(container_entrypoint.main(["--setup-chat", "--telegram-env", str(source)]), 0)
        process.assert_not_called()
        self.assertNotIn("--setup-chat", setup.call_args.args[0])


if __name__ == "__main__":
    unittest.main()
