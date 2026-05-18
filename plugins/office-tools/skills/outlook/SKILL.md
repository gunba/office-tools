---
description: Read-only Outlook mail and meeting search through the Rust office-tools WinCOM command. Use for inbox/sent/received/all-folder searches, saving matching .msg files, and saving attachments. This skill never sends mail.
---

# Outlook Mail Tool

Use on Windows with Outlook installed and authenticated:

```bash
"%LOCALAPPDATA%\Temp\office-tools\office-tools.exe" outlook [options]
```

Local development:

```bash
cargo run -- outlook [options]
```

## Commands

```bash
# Non-destructive COM availability check for Excel, Word, and Outlook.
office-tools doctor

# Recent unread inbox mail.
office-tools outlook --hours 24 --count 20

# Search received mail, including read items.
office-tools outlook --folder received --days 7 --include-read --search "invoice"

# Search sent items by recipient and subject.
office-tools outlook --folder sent --include-read --to "Tracy" --subject "Rinfo" --days 30

# Return full bodies and save matching attachments or .msg files.
office-tools outlook --sender "John Smith" --days 30 --include-read --full-body
office-tools outlook --search "engagement letter" --save-attachments C:\Temp\attachments
office-tools outlook --search "engagement letter" --save-msg C:\Temp\msg
```

Output is JSON with `emails`, `meeting_requests`, `saved_attachments`, and
`saved_msg`.

This tool exposes no send, reply, forward, delete, or move operations.

For full Windows COM runtime validation from a repo checkout, run:

```powershell
.\scripts\windows-wincom-smoke.ps1
```
