# Documentation

The canonical project README lives at the repository root: [../README.md](../README.md).
GitHub renders the root copy on the repository page, so it is the one to edit.

This folder holds the rest of the documentation:

| File | Contents |
|---|---|
| [AGENTS.md](AGENTS.md) | The rule set for anyone (human or agent) changing this project. |
| [CHANGELOG.md](CHANGELOG.md) | Version history, newest first, with timestamps. |
| `screenshots/` | PNG exports of `START HERE`, the Dashboard, and a manager tab, produced by `scripts\snapshot-report.ps1`. |
| `aoti-logo.png`, `aoti-logo-2023.svg` | Logo art used by the report Dashboard. |

Regenerate the screenshots after a layout change:

```powershell
.\scripts\snapshot-report.ps1 -Path "PM_Requests_Master_Template.xlsm" -Sheets "START HERE" -Prefix master
.\scripts\snapshot-report.ps1 -Path "Reports\PM Requests 2026-11-13 (11).xlsx" -Sheets "Dashboard","Alice Fontaine" -Prefix report
```