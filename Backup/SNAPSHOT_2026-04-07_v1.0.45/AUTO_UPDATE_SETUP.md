# Auto-Update Configuration

Gantech Efterkalk supports automatic updates via GitHub Releases.

## 1) Setup GitHub Repository

If you don't have one yet:
- Create a GitHub account (free)
- Create a repository named `efterkalk`
- Push your code there

## 2) Configure package.json

In your package.json, update the publish section:

```json
"publish": {
  "provider": "github",
  "owner": "YOUR_GITHUB_USERNAME",
  "repo": "efterkalk"
}
```

Replace `YOUR_GITHUB_USERNAME` with your actual GitHub username.

## 3) Increment version and build

When you make changes and want to release:

```
npm version patch
npm run build:win
```

This updates package.json version and builds the Windows installer.

Example version progression:
- 1.0.0 → 1.0.1 (patch = bug fix)
- 1.0.1 → 1.1.0 (minor = new feature)
- 1.1.0 → 2.0.0 (major = breaking change)

## 4) Create GitHub Release

1. Go to your GitHub repo
2. Click "Releases" → "Draft a new release"
3. Tag version: `v1.0.0` (matches your package.json version)
4. Title: `Gantech Efterkalk v1.0.0`
5. Upload the `.exe` file from `dist/` folder
6. Click "Publish release"

## 5) How updates work

When a user starts the app:
- App checks GitHub Releases for new version
- If found and newer than current, downloads in background
- When ready, shows "Opdatering tilgængelig" dialog
- User can restart now or later
- On restart, installer applies update automatically

## 6) Generate GitHub token (optional, for automation)

If you want to publish releases via CI/CD in future:

1. GitHub Settings → Developer settings → Personal access tokens
2. Generate token with `repo` scope
3. Keep it safe (like a password)

For now, manual release is fine for factory deployment.

## Notes

- First release: version must be `1.0.0` or higher
- Always increment version before build
- Tag format must be `v{version}` (e.g., `v1.0.1`)
- Users get notified automatically when new release is available
