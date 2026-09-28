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

Vite writes the static site to `dist/`. GitHub Actions builds and deploys `dist/` to GitHub Pages when changes are pushed to `main`. The local Nginx site serves the production build at `http://netentra.citcoms.up.ac.th/`.

## Local server deployment

A push to `main` also runs `.github/workflows/deploy.yml`. It verifies the Vite build, signs a webhook request with the GitHub Actions secret `DEPLOY_WEBHOOK_SECRET`, and sends it to `https://netentra.citcoms.up.ac.th/deploy`. Nginx forwards the request to the user service `p2027-webhook.service`. The service checks the HMAC signature, repository, branch, and timestamp, then runs `/home/piza/.local/bin/p2027-sync` to fetch the latest `main`, build it, and publish it in `/var/www/myapp/P2027_1_EntraID`. There is no periodic GitHub polling.

The webhook secret is stored on the server at `/home/piza/.local/state/p2027-webhook-secret` and must never be committed. The workflow pins the endpoint certificate in `.github/p2027-webhook.crt`. When replacing that certificate, update the pinned copy and reload the webhook service and Nginx. To inspect deployments on the server, run `journalctl --user -u p2027-webhook.service`.