# Deployment Form Generator

Generates deployment reports from PHub issues. It can run as a command-line tool or as the backend for `ccci-support`.

## Run the API for local development

1. Copy `.env.example` to `.env` and set `TOKEN` to your PHub API token. Keep `.env` private and out of version control.
2. Start the API from this folder:

   ```sh
   npm run dev
   ```

   It listens on `http://localhost:3001`. The browser UI sends the selected application and tag to `POST /api/deployment-report`; the API responds with an Excel file. `GET /api/health` can be used to check that the service is running.

3. In another terminal, start the `ccci-support` Vite app with `npm run dev`. Its development server proxies `/api` requests to port 3001.

For a deployed setup, set `CCC_SUPPORT_ORIGIN` to the exact origin serving the support app and expose the API only through a trusted, authenticated deployment environment. CORS is not authentication.

## Deploy the backend to Vercel

This project includes a Vercel Function at `api/deployment-report.js`. Import the `deployment-form-generator` repository into Vercel (or set the project Root Directory to `deployment-form-generator` if using a monorepo), then add these Environment Variables in Vercel:

- `TOKEN`: your PHub API token
- `CCC_SUPPORT_ORIGIN`: the full deployed `ccci-support` origin, such as `https://your-support-app.vercel.app`

Deploy the backend and note its URL. In the `ccci-support` Vercel project, add `VITE_API_BASE_URL` with that backend URL, then redeploy the frontend. The UI will call `<backend-url>/api/deployment-report`.

Keep the backend and report data behind an appropriate access control. CORS only controls browser access; it does not authenticate callers. The local `npm run dev` backend remains available for development.

The interactive CLI remains available with `npm run cli`.
