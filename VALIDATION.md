# Release validation — Payroll Desk 1.0.0

Verified on this Windows 11 x64 development computer, 30 September 2026.

| Check | Result |
|---|---|
| Next.js static export and TypeScript build | Passed |
| Python regression suite | 9 tests passed |
| Legacy August payroll migration | Calculated rows match the original application; source database untouched |
| Monthly overrides / locked-month API protection | Passed |
| Token, host and origin protection | Passed |
| Staged workbook import and locked-period rejection | Passed |
| PDF, Excel and Word generation | Passed |
| Backup, restore and invalid-backup rejection | Passed |
| 400-employee month calculation | Under the 2-second test threshold |
| 400-page payslip batch | 400 pages; under 60 seconds; long-name sample visually reviewed |
| Actual Qt desktop UI checks | 10 passed: dashboard, payroll, search, editor, close, reports, dark theme, settings, scales, native save bridge |
| High-DPI rendering | 200% scaling inspected; compact sidebar layout added |
| Packaged executable startup | Passed with PATH restricted to Windows directories (no Python/Node.js on PATH) |
| Installer execution | Exit 0; installed under the current user's LocalAppData Programs folder |
| Installed application | Rendered 400 synthetic employees; native bridge ready |
| Second launch / shutdown | Second launch exited 0; main app exited 0; no remaining application/WebEngine processes observed |
| Distribution privacy | No payroll database or salary workbook included in the installer |

Installer: `dist/PayrollDesk-Setup.exe` (approximately 132.7 MiB).

SHA-256: `6B313F0E9064EBCE89A209E8356F68166B9F0D0CF10F2C04AD20DE6F1721211D`

Screenshots and machine-readable UI results are under `artifacts/`. These are local verification artifacts, not part of the installer. The native file-dialog return was mocked for the automated save-bridge check; the Qt bridge and backup API were exercised.

## Not yet verified

- An independently clean Windows 10/11 computer or VM without Python/Node.js installed. Restricted-PATH execution on the build computer is useful evidence but is not that acceptance test.
- Real SMTP delivery. No employee emails were sent during development or verification.
- All printer models and display hardware combinations.

The tests currently emit one Starlette warning about a future test-client transport change; all assertions pass.
