# Setup

## 1. Backend (Google Apps Script)
1. Go to script.google.com -> New project. Paste the contents of `Code.gs`.
2. **Set the admin password:** Project Settings (gear icon) -> Script Properties -> Add script property
   - Property: `ADMIN_PASSWORD`
   - Value: your password (pick something strong; it is NOT stored in the code)
3. Pick `setup` in the function dropdown and click Run. Approve the permissions.
   This creates the "Slot Signup Data" spreadsheet with the Rooms and Bookings tabs.
   (Check the Execution log to confirm it did not warn about a missing password.)

## 2. Deploy
1. Deploy -> New deployment -> type: Web app.
2. Execute as: **Me**. Who has access: **Anyone**.
3. Deploy and copy the Web app URL (ends in `/exec`).

After any later change to `Code.gs`: Deploy -> Manage deployments -> edit (pencil) ->
Version: **New version** -> Deploy. The URL stays the same.

## 3. Front end
1. Paste the Web app URL into `API_URL` in `config.js`.
2. Host the files (`index.html`, `admin.html`, `config.js`, `style.css`) anywhere static,
   e.g. GitHub Pages or Netlify.

## 4. Use it
- Open `admin.html`, sign in, create an event, and share the link it shows ("Copy link").
- Participants enter name and a 10-digit mobile number, then pick a slot. The same number
  lets them find or change their slot later.
- "Export Excel" downloads one sheet with a row per booking.
