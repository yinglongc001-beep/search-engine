# Multiple VIN Search website

A local website version of the Tkinter VIN search. It reports vehicles in port stock, matching delivery records for VINs not present in port stock, and VINs with no match. Results can be downloaded as an Excel workbook with the same three sheets as the desktop app.

The app requires registration and sign-in. The administrator can create accounts, disable or enable them, change roles, delete accounts, and reset passwords. Passwords are hashed; the administrator cannot view a user's existing password. The administrator uploads one shared report ZIP; all signed-in users search the same dataset.

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

Put the 14 report files listed under “Required report files” on the page into one ZIP file. ZIP folder structure is fine; matching uses file names. The workbook sheets and columns must match the reports expected by the original script. The administrator selects the ZIP once and waits for confirmation that the shared data was saved. Users then paste VINs separated by newlines, commas, semicolons, or spaces and click **Search VINs**; searches reuse the saved port and delivery data.

Only the administrator can upload or replace the shared report set. The app saves the processed dataframes to `instance/shared_reports.pkl` locally, so users can search the same data and it survives app restarts. The uploaded source ZIP is not saved. Maximum ZIP size is 100 MB (500 MB after extraction). User accounts are stored in `instance/users.sqlite3` locally.

## Publish it online (Render)

This folder includes `render.yaml` for a Render web service with persistent storage for user accounts and the shared processed reports. After deployment, sign in as admin and upload the reports ZIP once. All users will then search the same data. Render will ask you to enter the admin password as a private environment variable during setup. `render.yaml` sets the admin username to `leo.chen`; do not add `.env`, the local `instance` folder, or report ZIPs to GitHub.

The persistent disk and paid web service in this deployment configuration incur hosting charges. Do not deploy until you have reviewed and accepted the price shown by Render. The processed shared data is kept on the persistent disk and loaded at app startup. See [Render's Flask deployment guide](https://render.com/docs/deploy-flask), [persistent disk documentation](https://render.com/docs/disks), and [pricing](https://render.com/pricing).
