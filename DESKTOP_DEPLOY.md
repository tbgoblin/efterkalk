# Gantech Efterkalk - Desktop deployment (Windows)

This project can now run as a desktop app (Electron) and be packaged as a Windows installer.

## 1) Install dependencies

Run from project folder:

npm install

## 2) Run desktop mode (development)

npm run desktop

This starts the internal server and opens the app window directly (no browser required).

## 3) Build Windows installer (.exe)

npm run build:win

Output folder:

- dist/

Typical artifact name:

- Gantech Efterkalk-Setup-<version>.exe

**Size note:** Build includes only production dependencies (express + mssql, ~15-20 MB packed). DevDependencies (electron/builder) excluded from installer.

## 3b) Branding (optional but recommended)

You can add branding assets in:

- build/

Suggested asset:

- build/icon.ico

If you add the icon, set this in package.json:

build.win.icon = "build/icon.ico"

## 4) Factory-friendly defaults included

- Native installer wizard (NSIS, not one-click)
- Per-machine install
- Desktop + Start Menu shortcuts
- Runs app after setup
- Optional auto-start at Windows login (enabled by default)
- Installer languages: Danish + English

## 5) Auto-start control

Auto-start is enabled by default.
To disable it for a launch, set environment variable before start:

Windows PowerShell:

$env:EFTERKALK_AUTO_START="0"; npm run desktop

Command Prompt:

set EFTERKALK_AUTO_START=0 && npm run desktop

## Notes for this project

- Backend is still your existing Node/Express app in server.js.
- Electron is only the desktop container.
- If SQL/native driver issues appear in packaged build, run:

npm run postinstall

or reinstall dependencies with a clean node_modules.

- For company distribution, code-signing is recommended (Windows SmartScreen trust).

## 6) Automatic updates

The app supports automatic updates via GitHub Releases.

See [AUTO_UPDATE_SETUP.md](AUTO_UPDATE_SETUP.md) for complete setup instructions.
