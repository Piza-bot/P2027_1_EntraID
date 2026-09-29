# P2027_1_EntraID
 
Sign-in app for Microsoft Entra ID. MSAL Browser reads the signed-in user's profile with delegated Microsoft Graph `User.Read`. After sign-in, the app provides Home, Dashboard, system tests, and contact navigation. The email test form collects recipient, subject, and message. A Node.js API validates the signed-in user and sends mail as `Piza@up.ac.th` with Microsoft Graph app-only access, restricted by Exchange Online RBAC for Applications.

## Local development

Requirements: Node.js 24 and npm.

```bash
npm install
cp .env.example .env
# Set the values in .env before starting the app.
npm run dev
```

Open the URL printed by Vite, normally `http://localhost:5173/`.

## Entra app registration and Exchange RBAC

1. In the Microsoft Entra admin center, create an app registration for **Accounts in this organizational directory only**.
2. Under **Authentication**, add these URLs as **Single-page application** redirect URIs:
	- `http://localhost:5173/redirect.html` for Vite development.
	- `https://netentra.citcoms.up.ac.th/redirect.html` for the production Nginx site.
	- `https://piza-bot.github.io/P2027_1_EntraID/redirect.html` for GitHub Pages.
3. Under **Expose an API**, use `api://<Application (client) ID>` as the Application ID URI and add a delegated scope named `access_as_user`.
4. Under **API permissions**, add Microsoft Graph **Delegated** permission `User.Read` and your app's delegated `access_as_user` scope. Grant admin consent if required by your tenant.
5. Set `ENTRA_TENANT_ID`, `ENTRA_CLIENT_ID`, and `ENTRA_CLIENT_SECRET` in `.env`. The client ID and tenant ID must also match `src/authConfig.js`. The client secret is used only by Node.js; never put it in frontend code or `VITE_*` variables.
6. Set `MAIL_FROM_ADDRESS=Piza@up.ac.th`.
7. In Exchange Online, assign the Node API's service principal the `Application Mail.Send` role with the Management Scope that includes `Piza@up.ac.th`. Verify it with `Test-ServicePrincipalAuthorization` for that mailbox. Do not also grant this app an unscoped Microsoft Entra **Application** `Mail.Send` permission, because grants are additive and would defeat the intended scope.

The browser app remains a public client and contains no secret. The secret belongs only to the Node API environment. The Node API checks the delegated `access_as_user` token, tenant, and client before it sends. Any signed-in user in the tenant who can access this app can send mail as `Piza@up.ac.th`; Exchange RBAC limits the mailbox the app can send as, not which signed-in users can call this API.

## Node API hosting

During development, Vite proxies `/api` to Node on port `3000`. In production, serve `dist/` with Nginx and reverse-proxy `/api/` to the Node service on the same origin, preserving the path. For example:

```nginx
location /api/ {
	proxy_pass http://127.0.0.1:3000;
	proxy_set_header Host $host;
	proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
}
```

Run the Node service with the required environment values from `.env` (or your process manager's secret store) and expose it only through the HTTPS Nginx site. GitHub Pages can host the static sign-in/profile page, but it cannot host this API; mail sending requires the Node service and a same-origin `/api` route.

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
