# Outlook Manager

A Docker Compose app for Outlook cleanup with a React frontend and Node backend.

## Features
- Single-user Azure AD login via MSAL
- Microsoft Graph message search and preview
- Bulk delete support
- Docker Compose deployment with Doppler secret handling

## Quick start
1. Copy `.env.example` to `.env`.
2. Set `VITE_AZURE_CLIENT_ID`, `VITE_AZURE_TENANT_ID`, and `VITE_API_BASE_URL`.
3. Run:
   ```bash
   docker compose up --build
   ```
4. Open the app at `http://localhost`.

## Deployment
This repo includes `.github/workflows/deploy.yml` for VPS deployment using SSH, SCP, and Doppler.
