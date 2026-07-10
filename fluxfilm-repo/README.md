# FluxFilm

Premium streaming-subscription reseller. **Google Sheets stays the master data
store.** This repo holds the code and drives a phased, safe migration of the
serving layer onto Hostinger for speed — without ever breaking what's live.

```
index.html → cache-server (Node) → api.php (Hostinger) → Apps Script → Google Sheets
```

## Repo layout

```
fluxfilm/
├── apps-script/        # Google Apps Script backend (.gs) — runs in Apps Script, NOT Hostinger.
│                       # Kept here for backup + version history. Deploy via the Apps Script editor / clasp.
├── hostinger/          # What actually deploys to Hostinger
│   ├── index.html          # the storefront (single-page app)
│   ├── api.php             # proxy to Apps Script (secrets loaded from config.php)
│   ├── config.example.php  # template — copy to config.php on the server
│   ├── config.php          # REAL secrets — gitignored, lives only on the server
│   └── cache-server/       # Milestone 1: Node.js caching layer in front of api.php
└── .gitignore
```

> **Apps Script vs Hostinger:** the `.gs` files are Google's language and only run
> inside Google Apps Script bound to the Sheet. Hostinger cannot execute them, so
> they are **not** deployed there — they live in `apps-script/` as a safe,
> versioned backup while the migration proceeds.

## 🔐 Security — do this before/at first push

1. `hostinger/config.php` holds live secrets and is **gitignored**. Never commit it.
2. **Rotate the keys** that were previously stored in plain text (treat as leaked):
   - the OpenAI API key
   - the `API_KEY` shared secret (update it in `config.php` **and** the Apps Script SETTINGS)
3. In `apps-script/TelegramAdmin.gs`, `tg_isAllowed_()` currently `return true` (allows
   anyone). Lock it to your admin chat ID before exposing the bot.

## Migration roadmap

- [x] **Milestone 1** — Node.js caching layer in front of `api.php` (`hostinger/cache-server/`).
- [ ] Move Gmail payment-checking to Hostinger.
- [ ] Add a MySQL mirror of the Sheet for fast reads + scheduled two-way sync.
- [ ] Move endpoints off Apps Script one at a time; retire it only once stable.

## Deploy notes

- **api.php** → upload `hostinger/` to your web root; create `config.php` from the
  template with real values.
- **cache-server** → see `hostinger/cache-server/README.md`.
- **Apps Script** → push `apps-script/` via the Apps Script editor or
  [clasp](https://github.com/google/clasp).
