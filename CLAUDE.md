# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Static wedding site for Eyla & Mauricio (June 6, 2026, Pereira, Colombia). Served from GitHub Pages at `maoaiz.github.io/wedding/` (remote: `maoaiz/wedding`). The backend is a Google Apps Script Web App bound to a private Google Sheet — there is no build step, no package manager, no tests.

UI copy is in Spanish. HTML entities (`&iacute;`, `&aacute;`, `é`…) are used instead of raw Unicode because some of this text is string-built from Apps Script and passes through JSONP.

## Running / developing

- Open any `.html` directly, or serve the directory: `python3 -m http.server 8000` then visit `http://localhost:8000/`.
- Test the RSVP flow with `?code=<codigo>` on `rsvp.html`, `index.html`, or `invitation/index.html`. The code is passed through from the landing page to the RSVP link and is case-insensitive / trimmed on the backend.
- Deploy = `git push`. GitHub Pages serves `main` at the project root.

## Architecture

Three entry pages, all standalone (each inlines its own CSS/JS, shares only `effects/*.css` and `images/`):

- `index.html` — save-the-date: hero photo, countdown, gallery slideshow, CTA to RSVP. Uses `images/` (save-the-date photo shoot).
- `invitation/index.html` — formal invitation with an interactive envelope-opening animation. Uses `invitation/preboda/` (preboda shoot) — do not swap these two photo sets, they are intentionally different.
- `rsvp.html` — RSVP form. Loads family/guest data by `code`, renders one card per guest, saves back.
- `admin.html` — private admin buttons that call Apps Script actions by name. The Sheet URL (`1rTGn5WOs_b3OFi1n3PVdrZZ8wA4Qm6ARYTQvcTfWgVw`) is hardcoded here.

### Frontend ↔ Apps Script (JSONP, not fetch)

All three dynamic pages talk to the same Apps Script Web App URL (`API_URL` constant, duplicated in `rsvp.html`, `invitation/index.html`, `admin.html`). They use **JSONP via injected `<script>` tags**, not `fetch` — this is deliberate: Apps Script `/exec` issues cross-origin redirects that `fetch` can't follow without CORS headers Google doesn't send. If you "modernize" this to `fetch`, it will break. The server-side counterpart is `respondGet()` in `docs/apps-script.js`, which wraps JSON in the `?callback=` name when present.

Actions are dispatched by `?action=...` query param: `get` (default — look up family by code), `save`, `generar_codigos`, `generar_mensajes`, `generar_mapa`, `setup_mesas`, `refresh_dropdown`, `crear_resumen`.

### Apps Script backend (`docs/apps-script.js`)

This file is a **mirror of code that lives inside the Sheet's Apps Script editor** — editing it here does nothing until it's pasted back into `Extensiones > Apps Script` in the Sheet and redeployed (`Implementar > Nueva implementacion` bumps the `/exec` URL — if you redeploy you must update `API_URL` in all three HTML files).

Key shape:

- Sheet `Invitados` columns: `Nombre | Apellido | Etiqueta | Telefono | Grupo | codigo | es_nino | confirmado | menu | notas | actualizado | visitas` (+ optional `mesa`). The code looks columns up by header name, so columns can be reordered but **header strings are the contract** — don't rename them casually.
- `confirmado` is stored in the Sheet as the strings `si` / `no` / `tal_vez` / `""`, but transported to the frontend as `true` / `false` / `'tal_vez'` / `''`. The conversion happens in `doGet` (read) and `saveResponses` (write).
- Every GET on a valid code increments the `visitas` column for every row in that group (engagement metric surfaced in `Resumen`).
- Every save triggers a notification email via `MailApp.sendEmail` to the couple. Email failures are swallowed so the save still succeeds.
- Functions come in pairs: the menu version (`generateCodes`, `setupMesas`, …) calls `SpreadsheetApp.getUi().alert(...)` and is bound to the `RSVP` custom menu via `onOpen`; the `...Remote` twin does the same work without UI calls and is what `admin.html` hits over HTTP. When changing logic, update **both**.
- `setup_mesas` is destructive (clears and rebuilds the `Mesas` tab + data validation). `refresh_dropdown` is the safe alternative when the list of tables changed but assignments should be kept.

Guests are grouped by the `Grupo` column — `generateCodes` assigns one shared 6-char code per group (alphabet excludes ambiguous chars like `0/O`, `1/l/i`). All members of a family share the same code and confirm together.

### Effects

`effects/*.css` are standalone CSS-only animations (petals, floral corners, bokeh, gradients). Each page picks one via `<link>`. The `bg-*.html` files at the repo root are previews / pickers for these effects, not production pages.

## Conventions worth preserving

- Don't introduce a build step, bundler, or framework — the site is intentionally zero-dependency static HTML.
- Keep `API_URL` in sync across `rsvp.html`, `invitation/index.html`, and `admin.html` when the Apps Script deployment is bumped.
- Keep `docs/apps-script.js` in sync with what's pasted in the Sheet. Treat changes here as requiring a manual redeploy step.
- `.gitignore` excludes `images/originals/` — high-res originals live there locally and should not be committed.
