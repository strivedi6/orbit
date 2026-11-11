# Orbit

Project management for hotel teams — projects, tasks, subtasks, deadlines, priorities, assignees, comments and files — with an iOS / macOS look.

**Open the app:** https://strivedi6.github.io/orbit/

## Install it like an app
- **iPhone / iPad:** open the link in Safari › Share › **Add to Home Screen**
- **Android:** open in Chrome › menu › **Install app**
- **Mac:** Safari › File › **Add to Dock**, or Chrome › install icon in the address bar
- **Windows:** Chrome or Edge › install icon in the address bar

## How it works
- `index.html` — the whole app (one file)
- `manifest.webmanifest`, `sw.js`, `icons/` — make it installable
- `backend/Code.gs` — Google Apps Script that stores data in the **Orbit Data** sheet and files in **Orbit Uploads** (Google Drive › Orbit)
- `backend/SETUP.md` — how the backend was set up

People sign up with a property code (e.g. DXBAI), choose their team, pick a username and 4-digit PIN, and wait for approval by the main admin or their department head.
