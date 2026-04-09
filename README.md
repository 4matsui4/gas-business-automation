# gas-business-automation

A Google Apps Script (GAS) toolkit for automating school submission management workflows — including Drive permission sync, deadline reminder emails, PDF generation, and Slack completion notifications.

---

## Features

| Feature | Description |
|---------|-------------|
| **Drive Permission Sync** | Automatically grants/revokes folder access based on email columns in the master sheet |
| **Grant Notification Email** | Sends an onboarding email to newly added recipients |
| **Deadline Reminder Email** | Sends reminder emails 4 days and 1 day before each form deadline |
| **PDF Auto-generation** | Converts submitted documents to PDF and stores them in Drive |
| **Completion Detection** | Detects when all forms are complete and notifies via Slack |
| **Edit Log** | Automatically records spreadsheet edits in a log sheet |

---

## System Overview

```
Master Spreadsheet (Master(完成版))
        │
        ▼ [Every 5 min trigger]
  syncAllPermissions
        │  Reads: Folder URL (col AG) + Emails (col AH~AJ)
        ▼
  Grant Drive folder access
        │  First-time recipients only
        ▼
  Send grant notification email (with Form 1 / Form 2 deadlines)
        │
        ▼ [Daily at 7:00 trigger]
  dailyReminder_
        │  d-4 / d-1 timing
        ▼
  Send reminder emails → log to AU/AV columns
        │
        ▼ [Every 15 min trigger]
  checkCompletionAndNotify_
        │  When col AD & AE both contain completion keywords
        ▼
  Auto-generate PDF → Slack notification
```

---

## Column Definitions (Master Sheet)

| Column | Content |
|--------|---------|
| D | School name |
| AB | Form 1 deadline date |
| AC | Form 2 deadline date |
| AD | Form 1 completion status |
| AE | Form 2 completion status |
| AF | Per-school submission book URL |
| AG | Parent Drive folder URL |
| AH~AJ | Contact email addresses |
| AK | Send status (完了 / エラー) |
| AL | Send metadata log |
| AU | Form 1 reminder send log |
| AV | Form 2 reminder send log |
| AO~AQ | Permission check results |

---

## Setup

### 1. Copy the script

Open your master Google Spreadsheet → **Extensions > Apps Script**, paste the contents of `submission-manager.gs`.

### 2. Replace placeholder values

Search for the following placeholders and replace them with your actual values:

| Placeholder | Location | Description |
|-------------|----------|-------------|
| `YOUR_MASTER_SPREADSHEET_ID` | Line ~100 | Your master spreadsheet ID (from the URL) |
| `YOUR_MASTER_SHEET_LINK` | Line ~50 | Full URL of the master sheet (for Slack buttons) |
| `YOUR_ORGANIZATION_NAME` | `SENDER.NAME` | Display name shown in sent emails |
| `info@your-domain.jp` | Reminder email body | Your actual inquiry email address |

### 3. Configure Slack (optional)

Set `SLACK.WEBHOOK_URL` to your Slack Incoming Webhook URL to enable completion notifications.

### 4. Enable required services

In the Apps Script editor → **Services**, enable:
- **Drive API** (for advanced permission checks including groups/domains)

### 5. Run initial setup

From the spreadsheet menu: **運用メニュー > 初期設定（認可）**

### 6. Create triggers

From the menu:
- **運用メニュー > 権限 > 権限付与の定期同期（5分）** — starts permission sync
- **運用メニュー > リマインド > 時間トリガー作成（毎日）** — starts daily reminders
- **運用メニュー > リマインド > 完了検知トリガー作成（15分）** — starts completion detection

---

## Configuration Reference

```javascript
// Reminder timing (days before deadline)
const REMINDER_OFFSETS = [4, 1];  // Sends at d-4 and d-1

// Completion keywords (partial match)
const COMPLETE_WORDS = ['完', '完了', '済', '〆', '終了', '提出済'];

// Stop words — rows with these statuses are skipped
const STOP_STATUSES = ['停止', 'キャンセル', '中止', '通知停止'];

// Permission role granted to recipients
CFG.PERMISSION_ROLE = 'editor';  // or 'viewer'
```

---

## File Structure

```
gas-business-automation/
├── submission-manager.gs   # Main script
└── README.md
```

---

## Notes

- All spreadsheet IDs and personal email addresses have been removed from this repository.
- Test functions (`testMailOnlyForMe`, `createReminderTestTrigger5min`) are included for development convenience — remove or disable them before deploying to production.
- This script uses `ScriptProperties` to persist state across trigger executions (managed permissions, reminder logs, completion flags).
