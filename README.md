# Chama WhatsApp Bot

Treasurer texts commands to a WhatsApp number; the bot reads and writes your Google Sheet (the Chama_Manager workbook).

## Setup (about 30-45 min)
1. **Google Sheet**: upload Chama_Manager.xlsx to Google Drive > open with Google Sheets > File > Save as Google Sheets. Copy the Sheet ID from the URL.
2. **Service account**: Google Cloud Console > new project > enable "Google Sheets API" > Credentials > Service account > Keys > JSON. Save as `service_account.json`. Share the sheet with the service account's email as Editor.
3. **WhatsApp provider** (pick one):
   - Twilio (easiest to test): WhatsApp Sandbox > set "When a message comes in" to `https://YOUR-URL/whatsapp` (POST). Members/officials must join the sandbox first. For production you need a Twilio WhatsApp sender.
   - Meta Cloud API: create app > WhatsApp > set webhook `https://YOUR-URL/whatsapp`, verify token = META_VERIFY_TOKEN, subscribe to "messages".
4. **Deploy** `app.py` to any host with HTTPS (Render, Railway, Fly.io...). Copy `.env.example` to environment variables. Start command: `gunicorn -w 1 --threads 4 app:app` (ONE worker: pending confirmations live in memory).
5. Set ADMIN_NUMBERS to the treasurer/officials. Everyone else can only send `bal` for their own balance (matched by the phone on the Members sheet).

## Contribution reports
Admins can request `daily` (today), `daily yesterday`, `daily 14 nov`, or `daily 2026-11-14`. Use `monthly` for the current month, `monthly last`, `monthly oct`, or `monthly 2026-10`. These reports use the dates and categories in the Transactions sheet. The monthly report compares savings received with the active members' monthly savings amounts and lists active members with no savings payment that month.

`group [month]` returns a formatted group statement with that month's collections, member balances, the current group position, and payment/treasurer details. Admins can forward the reply to their WhatsApp group; the report includes individual member balances, so share it only with the intended group. Long replies are split into messages of at most 1,500 characters.

## Safety built in
Writes need YES confirmation; `undo` removes your last entry; webhook signatures are verified; retried messages are not double-processed; only admin numbers can write.

## Limits
- Reads Summary/Loans/Dashboard computed values, so keep the sheet layout unchanged.
- Sending a statement *to a member* proactively (reminders) needs WhatsApp-approved templates; not included. Members can message the bot themselves within the normal 24-hour window.
- Test locally: `python test_local.py` (uses the .xlsx and LibreOffice).
