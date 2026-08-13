# Outlook Manager

A web app for managing Outlook mail via Microsoft Graph. Search, preview, move, delete, and manage inbox rules — all through a React frontend with a Node.js background job API.

## Features

- **Azure AD login** via MSAL (single-user)
- **Advanced message search** with filters:
  - Sender contains / does not contain (comma-separated multiple values)
  - Subject contains / does not contain
  - Body contains
  - Folder, date range, read status, attachments, importance
- **Message preview** with pagination (50 per page)
- **Bulk actions** on selected messages:
  - Delete selected
  - Move selected to folder
  - Delete all matching (background job)
- **Background jobs** for large operations (count matching messages, delete all matching)
- **Outlook inbox rules** management:
  - View all rules with full condition/action summaries
  - Create rules with multiple conditions and actions
  - Edit existing rules
  - Toggle enabled/disabled
  - Delete rules
- **Docker Compose** deployment with optional Doppler secret handling
- **GitHub Actions** workflow for VPS deployment

## Architecture

```
┌─────────────┐     ┌─────────────┐     ┌─────────────────┐
│   Browser   │────▶│   Nginx     │────▶│  web (React)    │
│             │◀────│  (reverse   │◀────│  static build   │
└─────────────┘     │   proxy)    │     └─────────────────┘
       │            └─────────────┘
       │
       ▼
┌─────────────────┐     ┌─────────────────┐
│ Microsoft Graph │     │  api (Node.js)  │
│   (Outlook)     │     │  background jobs│
└─────────────────┘     └─────────────────┘
```

- **web/** — React + Vite frontend. Talks directly to Microsoft Graph for interactive operations (search, delete, move, rules). Built as static files served by `serve`.
- **api/** — Express backend. Handles background delete/count jobs. Stores job state in memory.

## Local development

1. Copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```

2. Fill in the values:
   - `VITE_AZURE_CLIENT_ID` — from your Azure AD app registration
   - `VITE_AZURE_TENANT_ID` — `common` for multi-tenant, or your tenant ID
   - `VITE_REDIRECT_URI` — e.g. `http://localhost` (must match Azure redirect URI exactly)
   - `VITE_API_BASE_URL` — e.g. `http://localhost:4000`
   - `PORT` — API port, defaults to `4000`

3. Register redirect URIs in Azure AD:
   - **Authentication** → **Add a platform** → **Single-page application**
   - Add your local URL (e.g. `http://localhost`, `http://localhost:52048`, etc.)
   - For production behind nginx/HTTPS, add the public domain (e.g. `https://mail.example.com`)

4. Add API permissions in Azure AD:
   - `User.Read`
   - `Mail.ReadWrite`
   - `MailboxSettings.ReadWrite` (for inbox rules)
   - Grant admin consent

5. Run locally:
   ```bash
   docker compose up --build
   ```

6. Open the app at `http://localhost` (or the port mapped by Docker).

## Production deployment (VPS)

This repo includes `.github/workflows/deploy.yml` for automated VPS deployment via SSH and Doppler.

### Doppler setup

1. Store secrets in Doppler:
   - `VITE_API_BASE_URL`
   - `VITE_AZURE_CLIENT_ID`
   - `VITE_AZURE_TENANT_ID`
   - `VITE_REDIRECT_URI`

2. Add GitHub repository secrets:
   - `VPS_HOST`
   - `VPS_USER`
   - `VPS_SSH_KEY`
   - `DOPPLER_TOKEN`

3. Trigger deployment manually or push to `main`/`master`.

### Nginx reverse proxy (recommended)

Instead of exposing the API port publicly, proxy `/api` through nginx:

```nginx
server {
    listen 443 ssl;
    server_name mail.example.com;

    location / {
        proxy_pass http://outlook-cleaner-web:80;
        proxy_set_header Host $host;
    }

    location /api/ {
        proxy_pass http://outlook-cleaner-api:4000/;
        proxy_set_header Host $host;
    }
}
```

Set `VITE_API_BASE_URL=/api` in Doppler so the frontend uses same-origin requests.

## Environment variables

| Variable | Description | Example |
|----------|-------------|---------|
| `PORT` | API server port | `4000` |
| `VITE_API_BASE_URL` | Backend URL (baked into build) | `http://localhost:4000` or `/api` |
| `VITE_AZURE_CLIENT_ID` | Azure AD application client ID | `a1b2c3d4-...` |
| `VITE_AZURE_TENANT_ID` | Azure AD tenant ID | `common` or `e5f6g7h8-...` |
| `VITE_REDIRECT_URI` | MSAL redirect URI (must match Azure) | `http://localhost` |

## Testing

```bash
cd api
node --test *.test.js
```

## Tech stack

- **Frontend**: React, TypeScript, Vite, MSAL React
- **Backend**: Node.js, Express
- **Deployment**: Docker Compose, GitHub Actions, Doppler
