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
	- `https://netentra.citcoms.up.ac.th/redirect.html` for the production Nginx site.
	- `https://piza-bot.github.io/P2027_1_EntraID/redirect.html` for GitHub Pages.
3. Under **API permissions**, add Microsoft Graph **Delegated** permission `User.Read`.
4. Copy the **Application (client) ID** and **Directory (tenant) ID** from the app overview into `src/authConfig.js`.

The browser app is a public client. Do not create or add a client secret.

## Build and hosting

```bash
npm run build
```

Vite writes the static site to `dist/`. GitHub Actions builds and deploys `dist/` to GitHub Pages when changes are pushed to `main`. The local Nginx site serves the production build at `https://netentra.citcoms.up.ac.th/`.

## Local server deployment

A push to `main` also runs `.github/workflows/deploy.yml`. It verifies the Vite build, signs a webhook request with the GitHub Actions secret `DEPLOY_WEBHOOK_SECRET`, and sends it to `https://netentra.citcoms.up.ac.th/deploy`. Nginx forwards the request to the user service `p2027-webhook.service`. The service checks the HMAC signature, repository, branch, and timestamp, then runs `/home/piza/.local/bin/p2027-sync` to fetch the latest `main`, build it, and publish it in `/var/www/myapp/P2027_1_EntraID`. There is no periodic GitHub polling.

The webhook secret is stored on the server at `/home/piza/.local/state/p2027-webhook-secret` and must never be committed. The webhook uses the same publicly trusted HTTPS certificate as the site. To inspect deployments on the server, run `journalctl --user -u p2027-webhook.service`.

## HTTPS certificate

The production site uses a Let’s Encrypt certificate for `netentra.citcoms.up.ac.th`. `acme.sh` renews it through the HTTP challenge at `/.well-known/acme-challenge/`. The Nginx config in `ops/p2027-nginx.conf` serves that path on port 80, redirects other HTTP requests to HTTPS, and serves the site and `/deploy` over HTTPS.

The certificate and private key are installed at `/home/piza/.local/state/p2027-site/fullchain.pem` and `/home/piza/.local/state/p2027-site/privkey.pem`. The private key is kept out of Git. `p2027-cert-renew.timer` checks daily at 03:00 UTC with a random delay of up to 30 minutes. When renewal succeeds, the `acme.sh` install hook runs `p2027-cert-reload`, which validates the certificate, tests the Nginx configuration, and reloads Nginx. The limited sudoers rule in `ops/p2027-sudoers-nginx-reload` permits only the two required Nginx commands.

The installed scripts live in `/home/piza/.local/bin/`; copies of the scripts, user systemd units, Nginx config, and sudoers rule are in `ops/`. Run `p2027-cert-setup` after the Nginx config and limited sudoers rule are installed to issue a certificate if needed, install it, and register the reload hook. The server needs the verified `acme.sh` client at `/home/piza/.local/share/p2027-acme-client/acme.sh`. To check the timer and most recent renewal run:

```bash
systemctl --user status p2027-cert-renew.timer
journalctl --user -u p2027-cert-renew.service -n 50 --no-pager
openssl x509 -in /home/piza/.local/state/p2027-site/fullchain.pem -noout -dates
```

The Microsoft Entra app registration must include `https://netentra.citcoms.up.ac.th/redirect.html` as a **Single-page application** redirect URI for sign-in on this domain.
