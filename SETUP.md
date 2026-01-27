# Machine Paperwork Expiry Tracker - Setup Guide

## What This Does

This Google Apps Script automatically:
- **Calculates days remaining** until each machine document expires
- **Color-codes statuses**: OK (green), WARNING (yellow), URGENT (orange), CRITICAL (red), EXPIRED (dark)
- **Sends email alerts** at exactly 30, 14, 7, and 1 day(s) before expiry
- **Sends a daily summary** email listing all documents expiring within 30 days, with renewal requirements
- **Optionally sends WhatsApp messages** via CallMeBot (free) or Twilio

---

## Step-by-Step Setup

### 1. Open Your Google Sheet

Open the Google Sheet where you want to track your machine paperwork (or create a new one).

### 2. Open the Script Editor

Go to **Extensions > Apps Script** in the menu bar.

### 3. Paste the Code

1. Delete any existing code in the editor.
2. Copy the entire contents of `Code.gs` from this repository.
3. Paste it into the Apps Script editor.
4. Click the **Save** icon (or Ctrl+S).

### 4. Run the Setup

1. In the function dropdown at the top of the editor, select **`setupTracker`**.
2. Click the **Run** button (play icon).
3. When prompted, click **Review Permissions** and authorize the script to access your spreadsheet and send emails.
4. The script will create three sheets:
   - **Machine Tracker** — your main data sheet with sample data
   - **Config** — settings for email, WhatsApp, and alert preferences
   - **Notification Log** — records every notification sent

### 5. Configure Your Email

1. Go to the **Config** sheet.
2. Enter your email address in cell **B2** (next to "Email Address").

### 6. Enter Your Machine Data

1. Go to the **Machine Tracker** sheet.
2. Replace the sample data with your actual machines and documents.
3. Fill in these columns for each row:
   - **Machine ID** — Equipment identifier (e.g., EQ-001)
   - **Machine Name** — Full name (e.g., CAT 320 Excavator)
   - **Document Type** — Mulkiya, Insurance, Operating License, Third Party Inspection, etc.
   - **Expiry Date** — The document's expiry date (dd/MM/yyyy format)
   - **Renewal Requirements** — What you need for renewal (optional but recommended)
4. Leave **Days Remaining**, **Status**, and **Last Notified** empty — the script fills these automatically.

### 7. Test It

Use the **Paperwork Tracker** menu (appears in the sheet menu bar after reloading):
1. Click **Paperwork Tracker > Run Daily Update Now** to calculate all statuses.
2. Click **Paperwork Tracker > Send Test Email** to verify email delivery.

### 8. Enable Daily Automation

Click **Paperwork Tracker > Create Daily Trigger**.

This creates a trigger that runs every day at approximately 7:00 AM and:
1. Updates all days remaining and color-coded statuses
2. Sends threshold alerts for documents at exactly 30/14/7/1 days
3. Sends a daily summary of everything expiring within 30 days

---

## WhatsApp Notifications (Optional)

### Option A: CallMeBot (Free)

1. Send this message from your WhatsApp to the CallMeBot number:
   - Open WhatsApp and send `I allow callmebot to send me messages` to **+34 644 71 83 99**
2. You will receive an API key in the reply.
3. Go to the **Config** sheet and enter:
   - **WhatsApp Phone** (B8): Your phone number with country code, e.g. `+971501234567`
   - **WhatsApp API Key** (B9): The API key you received
   - **WhatsApp Method** (B10): `callmebot`

### Option B: Twilio (Paid, More Reliable)

1. Sign up at [twilio.com](https://www.twilio.com/).
2. Set up a WhatsApp sandbox or approved sender.
3. Go to the **Config** sheet and enter:
   - **WhatsApp Phone** (B8): Your phone number with country code
   - **WhatsApp Method** (B10): `twilio`
   - **Twilio Account SID** (B11): Your Twilio Account SID
   - **Twilio Auth Token** (B12): Your Twilio Auth Token
   - **Twilio From Number** (B13): Your Twilio WhatsApp sender number (e.g. `whatsapp:+14155238886`)

---

## Sheet Structure

### Machine Tracker Sheet

| Column | Name | Description |
|--------|------|-------------|
| A | Machine ID | Equipment identifier |
| B | Machine Name | Full equipment name |
| C | Document Type | Mulkiya, Insurance, License, etc. |
| D | Expiry Date | When the document expires |
| E | Days Remaining | Auto-calculated each day |
| F | Status | OK / WARNING / URGENT / CRITICAL / EXPIRED |
| G | Last Notified | Date of last alert sent (prevents duplicates) |
| H | Renewal Requirements | What's needed for renewal |

### Status Colors

| Status | Days Left | Color |
|--------|-----------|-------|
| OK | 31+ days | Green |
| WARNING | 15-30 days | Yellow |
| URGENT | 8-14 days | Orange |
| CRITICAL | 1-7 days | Red |
| EXPIRED | 0 or less | Dark/Black |

---

## Alert Schedule

| Trigger | When | What |
|---------|------|------|
| 30-day alert | Exactly 30 days before expiry | Email + WhatsApp per document |
| 14-day alert | Exactly 14 days before expiry | Email + WhatsApp per document |
| 7-day alert | Exactly 7 days before expiry | Email + WhatsApp per document |
| 1-day alert | Exactly 1 day before expiry | Email + WhatsApp per document |
| Daily summary | Every morning | Email digest of all items within 30 days |

---

## Troubleshooting

- **No menu appearing?** Reload the spreadsheet. The "Paperwork Tracker" menu loads on open.
- **No emails?** Check Config B2 has your email. Check your spam folder. Run "Send Test Email" from the menu.
- **Authorization error?** Re-run `setupTracker` and approve permissions again.
- **Trigger not firing?** Go to Apps Script editor > Triggers (clock icon on left) and verify the trigger exists. Delete and recreate if needed.
- **WhatsApp not working?** Make sure you completed the CallMeBot activation step (sending the allow message). Double-check phone format includes country code with `+`.
