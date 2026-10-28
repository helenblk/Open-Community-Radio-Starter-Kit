# Open Community Radio Starter Kit

A zero-cost workflow automation toolkit for independent and community radio stations, built entirely on Google Workspace (Sheets, Calendar, Gmail and Apps Script). It replaces scattered email threads, chat groups and manual spreadsheets with a small, automated system your whole team can use from the browser.

You **do not** need to know how to code. The files are meant to be copied into your station's Google Workspace. The **Setup & Documentation Manual (PDF)** in the releases walks you through every step.

## What's included

| Module | What it does |
|---|---|
| **1 – Application System** | Public web form where DJs apply for free broadcast slots. Reads free slots from your Google Calendar, blocks the slot right away (pending, yellow), logs the application in a Google Sheet and notifies the team. One click on "Accept" confirms the slot (green) and emails the DJ. |
| **2 – DJ Schedule Checker** | Self-service page: DJs enter the email address they applied with and see their upcoming dates – without exposing anybody else's schedule. |
| **3 – SOS Alert System** | A menu inside the Applications sheet that sends one BCC email to your standby list when a slot becomes free at short notice. |
| **4 – Supporting Membership & SEPA** | Membership form with a veto phase for the board, a pre-filled A4 PDF (incl. SEPA mandate), statutes and privacy policy attached automatically. |
| **Mailing List** | Newsletter sign-ups from both forms land in one shared sheet (email, name, DJ alias, source, date) – without duplicates. |

## Repository structure

```
Open-Community-Radio-Starter-Kit/
├── 1_Application_System_and_SOS/   ← Modules 1 + 3 (one Apps Script project)
│   ├── Code.js
│   └── Index.html
├── 2_DJ_Schedule_Checker/          ← Module 2 (standalone Apps Script project)
│   ├── Code.js
│   └── Index.html
├── 3_Supporting_Membership/        ← Module 4 (one Apps Script project)
│   ├── Code.js
│   ├── Index.html
│   └── PdfTemplate.html
├── LICENSE
└── README.md
```

In the Apps Script editor the files are called `Code.gs`, `Index` and `PdfTemplate` – paste the content of `Code.js` into `Code.gs`.

## Quick start

1. Copy the three spreadsheet templates (**Applications**, **Supporting Members**, **Mailing List**) – the link is in the manual.
2. Open each spreadsheet → **Extensions → Apps Script** and paste the matching files. The DJ Schedule Checker is a standalone project (**Drive → New → Google Apps Script**).
3. Fill in the configuration block at the top of every `Code.js` (station name, calendar ID, spreadsheet IDs, email addresses). Replace `[Your Station Name]` in the HTML files.
4. Set the Apps Script time zone to your station's time zone (**Project Settings**).
5. **Deploy → New deployment → Web app**, execute as *Me*, access *Anyone*.
6. Embed the `/exec` URL on your website with an `<iframe>`.

## Updating the code later

Always use **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**.
This keeps the web app URL – and therefore the links on your website – unchanged. "New deployment" would create a second URL while the old one keeps running the old code.

## Security & privacy

The toolkit is built with data protection in mind:

- **Forgery-proof accept links:** the "Accept" buttons only carry a random token. All data (email address, slot, name) is read from the sheet, so nobody can misuse the public web app URL to send emails in your station's name.
- **No double bookings:** a slot is re-checked and locked when an application is submitted.
- **Escaped input:** applicant input can't inject links or formatting into admin emails or pages, and text starting with `=`, `+`, `-` or `@` is stored as plain text (no formula injection).
- **Server-side validation:** required fields, email format and IBAN checksum are checked by the script, not only in the browser.
- **Exact lookups:** the Schedule Checker only matches complete email addresses.

You are still responsible for your organisation's own GDPR/DSGVO compliance (privacy policy, access rights to the sheets, deleting old data). For SEPA direct debits you also need a creditor identifier (Gläubiger-ID) – see the manual. This toolkit is not legal advice.

## Beyond radio

While built for a radio station, this is essentially a universal booking system. By renaming the form fields, the same code can manage appointments for rehearsal rooms, studios, workshops or consultation hours.

## License

Released under the [MIT License](LICENSE) – free to use, adapt and share for your own community projects.

## Contact

Built by Helen Bolek. Found a bug or have a question? Please open an **Issue** on GitHub or write to `helenbolek@web.de`.
