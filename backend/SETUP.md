# Orbit — setup guide

Orbit runs as one HTML file (on GitHub Pages, like Hearts) with a Google Sheet as the database. Until you connect the Sheet, it runs in demo mode with sample data saved only on the device.

## 1. Create the backend (10 minutes)

1. Open **Orbit Data** in your Drive folder **Orbit** (already created; files go to **Orbit › Orbit Uploads**). `Code.gs` already points to both.
2. **Extensions › Apps Script**. Delete the sample code and paste all of `Code.gs`.
3. At the top, edit `ADMIN` — your username and **4-digit PIN**. Edit `SEED` if you want different properties or teams.
4. Choose **setup** in the function menu and press **Run**. Approve the permissions (Sheets + Drive). The Sheet now has tabs for properties, teams, users, projects, tasks, comments and activity.
5. **Deploy › New deployment › Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Copy the web app URL (ends in `/exec`).

## 2. Connect the app

1. Open `index.html` and find `API_URL: ''` near the top of the script.
2. Paste the URL between the quotes: `API_URL: 'https://script.google.com/macros/s/…/exec'`
3. Upload `index.html` to a new GitHub repository and turn on **Settings › Pages**. Optional: add a 180×180 `icon.png` for the home-screen icon.
4. Open the link on iPhone (Safari › Share › **Add to Home Screen**), Android (Chrome › **Install app**) or Mac (Safari › File › **Add to Dock**).

## 3. How people join

1. Share the link and the property code (e.g. `DXBAI`).
2. They tap **Create Account → property code → their team → name, username, 4-digit PIN**.
3. You (main admin) or that team's department head approve them under **Team › Requests** and set their role and job title.
4. They stay signed in on that device until they choose **Settings › Sign Out**.

## Roles

| Role | Can do |
|---|---|
| Main Admin | Everything, across every property and team. Switch teams from the sidebar. Add properties and teams in Settings. |
| Department Head | Approve and manage people in their own team, assign roles (member/head), reset PINs, delete any task or project in their team. |
| Team Member | Create projects, tasks and subtasks; assign, prioritise, set deadlines, comment and upload files — within their own team only. |

## Good to know

- Files upload to **Orbit › Orbit Uploads** in your Drive (max 10 MB each) and open by link.
- Five wrong PINs lock an account for 15 minutes. A department head can reset a PIN.
- When you change `Code.gs` later, use **Deploy › Manage deployments › Edit › New version** so the URL stays the same.
- A 4-digit PIN is convenient but not strong security. Keep confidential documents (contracts with rates, guest data) in your normal Hilton systems and link to them instead.
