# Raise PO – Outlook add-in

Adds a **Raise PO** button to emails in Outlook. It reads the PO request email and its attached
quote or invoice PDF, then opens Nexvia with the supplier, ref, job, COS and lines ready to go.

Works in new Outlook, classic Outlook for Windows, Outlook on the web and Outlook for Mac
(needs a version that supports Mailbox requirement set 1.8).

## Files

| File | What it is |
|---|---|
| `manifest.xml` | Tells Outlook about the add-in. You install this file. |
| `taskpane.html / .css / .js` | The side panel that opens when you click **Raise PO**. |
| `parsers.js` | Reads the email and the PDF (same readers as the Nexvia Quote → PO script). |
| `commands.html` | Required by Outlook, does nothing else. |
| `assets/` | Button icons. |

## 1. Host the files on GitHub Pages

1. Sign in to github.com and create a new repository called **raise-po-addin** (Public).
2. **Add file → Upload files**, drag in everything from this folder (including the `assets` folder), then **Commit changes**.
3. **Settings → Pages**: under *Build and deployment* choose **Deploy from a branch**, branch **main**, folder **/ (root)**, then **Save**.
4. After a minute the page shows your address, e.g. `https://yourname.github.io/raise-po-addin/`.
   Open `https://yourname.github.io/raise-po-addin/taskpane.html` to check it loads.

The repository only holds the add-in code. No email content or PDFs are uploaded anywhere –
everything is read inside Outlook and passed straight to Nexvia in the link.

## 2. Point the manifest at your address

1. On GitHub, open `manifest.xml` and click the pencil (edit).
2. Replace every `https://REPLACE-ME.github.io/raise-po-addin` with your address from step 1
   (no slash on the end). There are 8 of them.
3. **Commit changes**, then download the edited `manifest.xml` to your computer.

## 3. Install it in Outlook (just you)

1. Go to **https://aka.ms/olksideload** (opens Outlook on the web's add-ins window).
2. **My add-ins → Custom Addins → Add a custom add-in → Add from File…** and choose `manifest.xml`.
3. Accept the warning. The add-in also appears in Outlook desktop after a few minutes (restart Outlook if needed).

If **Add a custom add-in** isn't there, your organisation has turned it off – ask IT to deploy
`manifest.xml` from the Microsoft 365 admin centre (**Settings → Integrated apps → Upload custom apps**).

## 4. Nexvia side

Keep these two Tampermonkey scripts installed in the browser you use for Nexvia:

- **Nexvia - Quote to PO Lines** (v1.25) – the panel on the PO page.
- **Nexvia - PO requests from Outlook** (v1.1) – receives what the add-in sends.

Make sure that browser is your **default browser**, because Outlook opens links there.

## Using it

1. Open a PO request email and click **Raise PO** (on the ribbon, or under **…** / **Apps**).
2. Check the supplier, ref, job and COS, and the lines read from the PDF. Edit or remove lines if needed.
3. Click **Open in Nexvia**. Nexvia opens with a banner showing the request.
4. Start a new PO. When the PO page opens, the Quote → PO panel fills in the lines, Supplier Ref
   and Acct# (matched from COS). Check them, then **Add lines to Nexvia PO** and **Create**.
5. Click **Done** on the banner.

Pin the panel (pin icon at the top) to keep it open as you move between emails.

## Settings

The Nexvia address the add-in opens is under **Settings** at the bottom of the panel.
Default: `https://lennox.nexvia.com.au/staff/accounts/purchase-orders` – change it if yours is different.

## Updating

Change the files on GitHub and the add-in picks them up next time it opens.
Only re-install `manifest.xml` if you change the manifest itself.
