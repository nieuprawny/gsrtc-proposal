# GSRTC LED Screen — Proposal Generator

A small web app that lets you and your team generate customized PowerPoint and Excel proposals for GSRTC LED Screen advertising clients.

## What it does

1. Enter client name, mobile, email
2. Select one or more bus stand locations from 20 options
3. Click "Download PPT" → get a customized presentation (cover + selected locations only + benefits + T&C)
4. Click "Download Excel" → get a rate card with only the selected locations + auto-calculated totals

Both files get the client name in the filename so they're ready to send.

## Running locally

```bash
pip install -r requirements.txt
uvicorn app:app --reload
```

Then open http://localhost:8000

## Deploying to Render.com (recommended — free tier works)

### One-time setup

1. **Create a GitHub repo:**
   ```bash
   cd proposal_app
   git init
   git add .
   git commit -m "Initial commit"
   ```
   Create an empty repo on github.com, then:
   ```bash
   git remote add origin https://github.com/YOUR_USERNAME/gsrtc-proposal.git
   git push -u origin main
   ```

2. **Connect to Render:**
   - Go to https://render.com and sign up (free)
   - Click "New +" → "Web Service"
   - Connect your GitHub repo
   - Render auto-detects `render.yaml`. Click "Create Web Service"
   - Wait ~3 minutes for first deploy
   - You'll get a URL like `https://gsrtc-proposal.onrender.com`

3. **Share the URL with your team.** Bookmark it on your phone too.

### Important notes about the free tier

- **App sleeps after 15 min of no use.** First load after sleep takes ~30 seconds while it wakes up. After that it's instant.
- If this is annoying, upgrade to the $7/month tier — no sleep.
- Files are generated in-memory and downloaded directly. No data is stored on the server.

## Deploying to other platforms

The app is a standard FastAPI app — works anywhere that runs Python. Same code deploys to:
- **Railway** (railway.app) — $5/mo, no sleep
- **Fly.io** — generous free tier
- **Heroku** — paid only now
- **Your own VPS** — `uvicorn app:app --host 0.0.0.0 --port 80`

## Updating the template

If you want to change the PPT design, T&C, or pricing:

- **Template PPT**: replace `assets/template.pptx` with your new version. The slide order MUST stay the same (cover, summary, data, media report, then 2 slides per location in the same order, then benefits, then T&C). If you add/remove locations or change slide order, update the `LOCATION_SLIDE_MAP` in `generator.py`.

- **Rate card**: replace `assets/ratecard.xlsx`. The structure (which cells contain what) must match. If you add new locations, update `LOCATIONS_IN_ORDER` in `generator.py`.

After making changes, commit and push to GitHub — Render will auto-deploy.

## Files

- `app.py` — FastAPI web server (routes + HTML rendering)
- `generator.py` — PPT and Excel generation logic
- `templates/index.html` — frontend UI
- `assets/template.pptx` — your 46-slide master template
- `assets/ratecard.xlsx` — your 20-location pricing sheet
- `requirements.txt` — Python dependencies
- `render.yaml` — Render.com deploy config

## Troubleshooting

**"Module not found" on Render:** Make sure `requirements.txt` is committed and pushed.

**Slow first load:** Free tier cold start, normal. Pay $7/mo to eliminate it.

**Excel shows #VALUE! errors:** The template formulas reference specific rows. If you change the rate card structure, the formula-renumbering logic in `generator.py` may need updating.

**Want a logo or different colors:** Edit `templates/index.html` — change the header section and the `#C00000` color value to whatever you want.
