# Security notes

- Real bot tokens belong only in an untracked local `telegram.env` or a controlled secret store. The example file deliberately contains empty values.
- Never include tokens in issues, pull requests, source files, screenshots, shell commands or repository URLs. Do not paste complete Telegram request errors from another client if they contain the token in a URL.
- The application verifies HTTPS certificates and refuses redirects for Telegram API calls. It does not log tokens or raw API errors.
- Launchers mount settings/input/templates read-only and write reports only to the mounted output. A temporary owner-only settings copy inside Docker is deleted after use. Protect the original host file with chmod 600 on Linux or restrictive Windows ACLs.
- Runtime input/output and compiled images are excluded by `.gitignore`; the Docker context allows only named source/dependency files. Browser uploads still require manual review because .gitignore does not filter them.
- Use a Telegram chat you are authorized to send measurement information to. Treat reports and warnings as potentially sensitive operational data.
- CI uses read-only repository permissions, does not persist checkout credentials, and has no Telegram secrets. Network requests in tests are mocked.
- Docker administrators and a compromised host can access mounted settings or processes. These controls do not make a shared/untrusted machine safe for secrets.

If a bot token is exposed, revoke it through BotFather before investigating copies in commits, archives, logs or screenshots. Removing a token from the latest file alone does not remove it from Git history.

For a security report, use GitHub's private vulnerability reporting if the repository owner enables it. If no private channel is available, ask the owner for one without publishing secrets, private measurements or exploit details in a public issue.
