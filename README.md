# The Directory — Static Frontend for Business Card Registry

A refined editorial-style static page that displays your business card directory from Google Sheets.

## What you got

**`index.html`** — Single-file static page (HTML + CSS + JS embedded).
- Mobile-first responsive design
- Editorial magazine aesthetic (warm cream + terracotta accent + Fraunces display serif)
- Live search across all fields
- Detail drawer with LinkedIn + company point-by-point profile
- LinkedIn DP auto-derived from profile slug via `unavatar.io` (falls back to initials)
- Smooth animations, accessible, keyboard-friendly (ESC to close drawer)

**`BusinessCard_AppScript_V4_FINAL.js`** — Updated Apps Script with new JSON endpoint.

---

## Setup (one-time, 5 minutes)

### Step 1 — Update Apps Script

1. Open your Apps Script editor (Tools → Apps Script from the Sheet)
2. **Replace** the existing code with the new `BusinessCard_AppScript_V4_FINAL.js`
3. Click **Deploy → Manage deployments**
4. Edit the existing deployment → click the pencil icon → **Version: New version** → Deploy
5. **Critical:** Ensure "Who has access" is set to **Anyone** (not "Anyone within domain"), so GitHub Pages can fetch it without login.

### Step 2 — Verify JSON endpoint works

Open this URL in your browser:
```
https://script.google.com/a/macros/laserpowerinfra.com/s/AKfycbw9mmLORN7o7_IIu0YvnfikXmAvaZDq2dyzCCOkkmzMwm1eA2c4pKgtvmLlnubSlBA/exec?format=json
```
You should see raw JSON with all your cards. If you see HTML, the deployment hasn't picked up the new code — redeploy with a new version.

### Step 3 — Host `index.html` on GitHub Pages

1. Create a new GitHub repo (or use existing `visiting-card` repo)
2. Add `index.html` to the root
3. Go to **Settings → Pages**
4. Source: `main` branch, root folder → **Save**
5. Wait ~30 seconds, GitHub will give you a URL like `https://laserpowe.github.io/directory/`

That's it. The page will fetch from your Apps Script on every load.

---

## Customising

Open `index.html`. Top of the `<script>` section:

```javascript
const API_URL = "https://script.google.com/a/macros/laserpowerinfra.com/s/AKfycbw9mmLORN7o7_IIu0YvnfikXmAvaZDq2dyzCCOkkmzMwm1eA2c4pKgtvmLlnubSlBA/exec";
```

Replace if you ever redeploy and get a new URL.

### Theme colors

Top of `<style>` section, change CSS variables:
```css
--accent:       #C8553D;        /* terracotta */
--bg:           #FAF7F2;         /* warm cream */
```

Want dark mode? Swap `--bg`, `--ink`, `--line` etc. for darker values.

---

## How the data flows

```
[Sheet]  →  [Apps Script doGet?format=json]  →  [GitHub Pages index.html]  →  [Browser renders cards]
```

- Apps Script reads the sheet and serves JSON
- Static page fetches JSON on load
- Cards render with all enrichment data (LinkedIn URL, position bullets, company profile, etc.)

The Apps Script also still serves the legacy HTML view at the same URL (without `?format=json`), so existing bookmarks keep working.

---

## Troubleshooting

**"Couldn't reach the directory API"**
- Open the API URL directly in browser. If you see "Sign in" page, your deployment is not public — redeploy with "Who has access: Anyone".

**LinkedIn photos not showing**
- This is expected for private profiles. `unavatar.io` only works for public LinkedIn profiles. Cards fall back to initials in a stylish gradient circle.

**Cards show "Enrichment issue" text**
- Run "🔍 Enrich Selected Row" from the Apps Script menu for those rows. The text will be replaced with the proper bullet-point profile.

**Mobile view broken**
- Hard-refresh the page (Ctrl+Shift+R) to clear cached CSS.
