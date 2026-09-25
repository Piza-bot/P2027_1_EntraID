# P2027_1_EntraID
 
Single-page sign-in app for Microsoft Entra ID. It uses MSAL Browser and the delegated Microsoft Graph `User.Read` permission to show the signed-in user's `/me` profile.

## Local development

Requirements: Node.js 24 and npm.

```bash
npm install
npm run dev
```

Open the URL printed by Vite, normally `http://localhost:5173/`.

## Entra app registration

1. In the Microsoft Entra admin center, create an app registration for **Accounts in this organizational directory only**.
2. Under **Authentication**, add these URLs as **Single-page application** redirect URIs:
	- `http://localhost:5173/redirect.html` for Vite development.
	- `http://localhost/redirect.html` for the local Nginx site.
	- `https://piza-bot.github.io/P2027_1_EntraID/redirect.html` for GitHub Pages.
3. Under **API permissions**, add Microsoft Graph **Delegated** permission `User.Read`.
4. Copy the **Application (client) ID** and **Directory (tenant) ID** from the app overview into `src/authConfig.js`.

The browser app is a public client. Do not create or add a client secret.

## Build and hosting

```bash
npm run build
```

Vite writes the static site to `dist/`. For Nginx, set the site's document root to `/home/piza/projects/P2027_1_EntraID/dist`, then validate and reload Nginx. GitHub Actions builds and deploys `dist/` to GitHub Pages when changes are pushed to `main`.