# FluxFilm Cache Server (Milestone 1)

A tiny Node.js layer that sits **in front of** `api.php` so read-only calls feel instant.

```
index.html  ->  cache-server  ->  api.php  ->  Apps Script  ->  Google Sheets
```

Only three read-only actions are cached in memory: `getBootstrap`, `getStockLevels`,
`getTrendingItems`. Everything else (orders, payments, OTP, admin, etc.) is passed
straight through and **never cached**. If the upstream ever errors, cached routes
fall back to the last good copy, and non-cached routes just return the error — so
this layer can never "break what's live."

## Run locally

```bash
cd hostinger/cache-server
cp .env.example .env      # then edit .env with real values
npm install
npm start
```

Test it:

```bash
curl http://localhost:8080/health
curl -X POST http://localhost:8080/api -H "Content-Type: application/json" \
  -d '{"action":"getBootstrap"}' -i    # look for the X-Cache: HIT/MISS header
```

## Endpoints

| Method | Path                     | Purpose                                  |
|--------|--------------------------|------------------------------------------|
| POST   | `/api`                   | Main proxy (frontend posts here)         |
| GET    | `/health`                | Health check + list of cached keys       |
| GET    | `/clearcache?key=...`    | Wipe the in-memory cache                  |

## Deploy on Hostinger

1. Create a **Node.js app** in hPanel, point it at this `cache-server` folder.
2. Set the environment variables from `.env.example` in the hPanel UI.
3. Start command: `npm start` (entry `server.js`).
4. Once verified, point the frontend's API base to this server's `/api` URL.
   Until you do that switch, the old `api.php` path keeps working exactly as before.

## Frontend switch (only when ready)

In `index.html`, change the API base URL from `.../api.php` to the cache
server's `/api` endpoint. Roll back instantly by pointing it back to `api.php`.
