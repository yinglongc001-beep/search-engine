# Multiple VIN Search website

A local website version of the Tkinter VIN search. It reports vehicles in port stock, matching delivery records for VINs not present in port stock, and VINs with no match. Results can be downloaded as an Excel workbook with the same three sheets as the desktop app.

The app now requires registration and sign-in. The administrator can create accounts, disable or enable them, change roles, delete accounts, and reset passwords. Passwords are hashed; the administrator cannot view a user's existing password. A user's uploaded report data is isolated to that browser session.

## Run it

Install Python 3.10 or newer, then double-click `start-website.bat`. On the first run it creates a private Python environment, installs the requirements, and opens the website in your browser. Keep the command window open while using the site; close it to stop the server.

You can also start it manually in a terminal in this folder:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000. The app runs locally; report data is uploaded only to the local server process.

## Prepare report ZIP

Put the 14 report files listed under “Required report files” on the page into one ZIP file. ZIP folder structure is fine; matching uses file names. The workbook sheets and columns must match the reports expected by the original script. Select the ZIP and wait for the page to say the reports are loaded. The app builds the port and delivery data once and keeps them in memory, so subsequent VIN searches reuse them. Paste VINs separated by newlines, commas, semicolons, or spaces, and click **Search VINs**.

Each signed-in browser session's processed reports are kept separately in memory for up to two hours after its last activity. The app does not save uploaded reports to disk. Maximum ZIP size is 100 MB (500 MB after extraction), and up to six report sets can be cached at a time. User accounts are stored in `instance/users.sqlite3` locally.

## Publish it online (Render)

This folder includes `render.yaml` for a Render web service with persistent storage for user accounts. To publish it, put this project in a GitHub repository, then in Render choose **New > Blueprint** and connect that repository. Render will ask you to enter the admin password as a private environment variable during setup. `render.yaml` sets the admin username to `leo.chen`; do not add `.env`, the local `instance` folder, or report ZIPs to GitHub. Each visitor uploads their own reports on the page.

The persistent disk and paid web service in this deployment configuration incur hosting charges. Do not deploy until you have reviewed and accepted the price shown by Render. Uploaded report data is kept in memory and clears if the server restarts; users can upload the ZIP again. See [Render's Flask deployment guide](https://render.com/docs/deploy-flask), [persistent disk documentation](https://render.com/docs/disks), and [pricing](https://render.com/pricing).
