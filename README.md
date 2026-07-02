# KE&G Employee Choice Awards — Netlify Deploy

## What's in here
- `index.html` — loads React/Babel/XLSX from CDN (no build step, same pattern as your other single-file tools)
- `app.jsx` — the full voting app (voting, admin, categories, roster import, everything you've built)
- `netlify/functions/storage.js` — backend that replaces the roster/categories/votes storage, using Netlify Blobs
- `node_modules/`, `package.json`, `package-lock.json` — the one dependency the storage function needs (`@netlify/blobs`), pre-installed so drag-and-drop deploy works without running `npm install` yourself

## Deploy (drag-and-drop, no GitHub needed)
1. Go to https://app.netlify.com/drop
2. Drag this whole folder (or unzip the .zip and drag the `eca-netlify` folder) onto the page
3. Netlify gives you a live URL immediately — something like `random-name-123.netlify.app`
4. Open it, go to Admin, and set your passcode like before

## Renaming the site
In Netlify: Site settings → Change site name, to get a cleaner URL (e.g. `keg-awards.netlify.app`).

## Redeploying after I make more changes
Same as now: drag the updated folder onto https://app.netlify.com/drop again — it deploys as a new version of the same site (as long as you drop it onto the existing site's deploy page, not a fresh one).

## Data
Everything (roster, categories, votes, passcode) now lives in Netlify Blobs — not in the Claude artifact anymore. That means:
- It works for anyone who visits the URL, no Claude account needed
- It survives independent of this chat — I can't reach into it from here anymore to inspect or fix data directly; changes now happen through the Admin panel on the live site itself
- If you want a fresh event next year, use the "Start Fresh" button in Admin → Import, same as before

## Passcode
Same simple passcode approach as before — set on first Admin visit, stored alongside the rest of the data. It's a casual-visitor deterrent, not real security, same tradeoff as when this lived in Claude.
