# Clinique laser — patient management on Google Sheets

A small patient-management app for a laser clinic. A Google Sheet is the database; an HTML app (opened from the sheet's menu, or as a web page) is where everything is done: registering patients, viewing and editing profiles, the appointment calendar and the statistics.

The interface is in French.

## What's in this repository

| File | What it is |
|---|---|
| `Clinique_Laser_v5.xlsx` | The spreadsheet template. Import it into Google Drive to create the clinic's sheet. |
| `Script_Clinique_Laser_v5.gs` | The Apps Script code (server side). Goes into the sheet's `Code.gs`. |
| `App.html` | The app's interface. Goes into an Apps Script HTML file named **`App`**. |

### The spreadsheet

Only two tabs:

- **Registre patients** — the database. One row = one appointment/visit. Do not reorder or rename the columns: the app relies on their position.
- **Liste des lazéristes** — the names offered in the app's "Lazériste" field. One name per row, **starting at row 5**.

| Col. | Header | Content |
|---|---|---|
| A | Terminé ? | Checkbox — ticked when the treatment is done |
| B | Date d'enregistrement | Set automatically when the patient is registered |
| C | Nom | Last name (required) |
| D | Prénom | First name (required) |
| E | Région / Wilaya | |
| F | Téléphone | Stored as text (keeps the leading 0) |
| G | Heure d'arrivée | HH:mm |
| H | Date du rendez-vous | dd/MM/yyyy |
| I | Heure du rendez-vous | HH:mm |
| J | Lazériste | |
| K | Paramètres / Type de laser | Free text |
| L | Séances | Whole number |
| M | Prix (DA) | Amount in Algerian dinars |
| N | Remarques | Free text |

Row colours: **green** row = treatment done; **red** checkbox cell = patient on the row, not done yet; every other filled row lightly shaded.

### The app

Four tabs:

- **Nouveau patient** — registration form with validation; the patient is written to the first empty row of the register.
- **Patients** — searchable, filterable list (À venir / Non terminés / Terminés / Sans RDV) and a profile for each patient: view all details, **edit**, **mark as done**, **delete** (with confirmation), other visits of the same patient, link to the row in the sheet.
- **Calendrier** — month calendar with the number of appointments per day; click a day to list its appointments, click an appointment to open the profile.
- **Statistiques** — income, appointments, unique patients, average price, sessions, completion rate, each compared with the previous period; charts over time, by lazériste, by region, by weekday, status breakdown and a list of past treatments not marked as done. Filters: period (today, this week, this month, last month, this year, all, custom dates), lazériste, and whether to count by appointment date or registration date.

Safety built in: before editing or deleting, the app checks the row still holds the same patient (in case the sheet changed meanwhile); writes are queued so two people saving at once don't overwrite each other; unsaved edits ask for confirmation before being discarded.

## Setup, step by step

### 1. Create the sheet

1. Open [Google Drive](https://drive.google.com) → **New → File upload** → choose `Clinique_Laser_v5.xlsx`.
2. Open the uploaded file with **Google Sheets**. If the title bar shows an `.xlsx` badge, use **File → Save as Google Sheets** and work in the new copy.
3. In **Liste des lazéristes**, type the lazéristes' names from cell **A5** downwards.

### 2. Add the script

1. In the sheet: **Extensions → Apps Script**.
2. Open `Code.gs`, delete its content, paste the full content of `Script_Clinique_Laser_v5.gs`.
3. Click **+** next to *Files* → **HTML**, name it exactly **`App`** (Apps Script adds `.html`), delete its content and paste the full content of `App.html`.
4. Click **Save** (💾).
5. Make sure there are no other `.gs` or `.html` files in the project (delete any leftovers via **⋮ → Delete**).

### 3. First run

1. Go back to the sheet and **reload the page**. A **Clinique laser** menu appears (it can take a few seconds).
2. Run **Clinique laser → Configurer le registre**.
3. Google asks for authorisation the first time: **Continue**, choose your account, then **Advanced → Go to (project name) → Allow**. The warning appears because the script is private to you, not published by Google.
4. When the "Configuration terminée" message appears, the register is ready: checkboxes in column A, colours, formats, frozen header, filter.

> Google Sheets cannot import checkboxes from an Excel file — this step is what creates them. It can be re-run at any time; it only re-applies formatting and never changes the data.

### 4. Use the app

**Clinique laser → Ouvrir l'application** opens the app in a window over the sheet.

### 5. Optional — use it as a web page

To use the app in a normal browser tab (or bookmark it on another computer):

1. In Apps Script: **Deploy → New deployment** → type **Web app**.
2. **Execute as:** *Me*. **Who has access:** *Only myself* (or your Google Workspace domain). **Do not choose "Anyone"** — this is patient data.
3. **Deploy**, then copy and bookmark the web app URL.

After changing the code later, update the same URL instead of creating a new one: **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**.

## Moving data from an older sheet

1. In the old sheet, select the data rows of **Registre patients**, columns **A to N** (not the header), and copy.
2. In the new sheet, click **A2** in **Registre patients** → **Edit → Paste special → Values only**.
3. Copy the lazéristes into the new **Liste des lazéristes** from **A5**.
4. Run **Clinique laser → Configurer le registre** again.

## Updating the code

1. Replace the content of `Code.gs` and/or `App` in Apps Script with the new versions, then **Save**.
2. Reload the sheet.
3. If you use the web app: **Deploy → Manage deployments → ✏️ → New version → Deploy**.

## Maintenance

- **Remove old scripts:** Apps Script → *Files* → **⋮ → Delete** on anything other than `Code.gs` and `App`. Old files defining `onOpen` / `onEdit` conflict with the current code.
- **Remove old deployments:** **Deploy → Manage deployments** → select → **Archive** (the URL stops working).
- **Triggers:** none are needed. The ⏰ *Triggers* page should be empty; delete anything listed there from older experiments.
- **Growing beyond 1,000 rows:** nothing to do. When the register is full, the script adds 50 rows with the same formatting, colours, checkboxes and filter, and keeps doing so. Google Sheets allows about 700,000 rows at 14 columns. The app reloads the whole register after each save, which stays fast up to a few thousand rows; beyond ~10,000, consider moving older years to an archive.
- **Undoing a deletion:** deleting a patient in the app removes the row. Recover it with **File → Version history** in the sheet.

## Troubleshooting

| Symptom | Fix |
|---|---|
| No **Clinique laser** menu | Reload the sheet; check `Code.gs` is saved and contains `onOpen`. |
| "Impossible de lire le registre" in the app | The tabs must be named exactly **Registre patients** and **Liste des lazéristes**. |
| Lazériste list is empty in the form | Names must start at **A5** in **Liste des lazéristes**; click ↻ in the app. |
| Column A shows TRUE/FALSE instead of checkboxes | Run **Clinique laser → Configurer le registre**. |
| "Cette fiche a changé dans le registre entre-temps" | The row was edited or deleted in the sheet while the app was open. Click ↻ and try again. |
| Web app shows an old version | Publish a new version: **Manage deployments → ✏️ → New version**. |

## Privacy

The sheet contains personal patient data. Share it only with authorised staff, keep the web app restricted to *Only myself* or your organisation, and avoid detailed medical information in the *Remarques* field.
