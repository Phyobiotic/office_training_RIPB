# Wiring this into your project

## Files

- `manifest.xml` → replace your project's root manifest, keep your own `<Id>`.
- `taskpane.html` → replace `src/taskpane/taskpane.html`. Self-contained
  (own `<style>`/`<script>`), so `taskpane.js`/`taskpane.css` are unused —
  safe to ignore or delete.
- `lessons/*.js` → new folder, `src/taskpane/lessons/`. Each file is one
  lesson, loaded dynamically at runtime (not bundled as a webpack entry).

## One-time setup: serve the lessons/ folder

This is the step that's easy to miss and will make lessons silently fail
to load if skipped. Webpack only serves files it's explicitly told about.

Put your lesson files at `assets/lessons/*.js` (source location) — that's
where the task pane requests them from at runtime
(`assets/lessons/<id>.js`), so the CopyWebpackPlugin `from` pattern needs
to point at that same source location, copying to the matching spot in the
build output:

```js
{ from: "assets/lessons/*.js", to: "assets/lessons/[name][ext]" }
```

The `from` is where the files live in your project; the `to` is where
they land in the served output — if you move the source folder, update
`from` to match, or copying silently does nothing and you'll get a 404
with no other clue why.

Restart `npm start` after any webpack config change (config edits aren't
picked up by hot reload).

## Adding a new lesson

1. Copy `lessons/xl-data.js` as your starting point.
2. Change the title, hook, brief, `setup()`, and `tasks` for both
   `tutorial` and `assignment`.
3. Save it as `lessons/<your-id>.js` in `src/taskpane/lessons/`.
4. Add one line to the `LESSONS` array in `taskpane.html`:
   `{ id:"<your-id>", title:"...", ready:true }`
5. That's it — no other file needs touching. This is deliberate: your
   friends can each own one lesson file without needing to understand or
   edit the shell.

## What to test

1. Open the pane — "Entering and Managing Data" and "Founders' Day" should
   be the only clickable rows.
2. Open "Entering and Managing Data" → **Tutorial** → **Load starter
   data** → work through the roster in Excel → **Check my work**.
3. Switch to **Assignment** → **Load starter data** → work through the CIP
   log → **Check my work**. This one's deliberately denser — 15 rows, four
   spelling variants to normalise, a non-adjacent selection.
4. Open "Founders' Day" and repeat — no Tutorial/Assignment toggle on this
   one since it's already the terminal assignment.

## Known rough edges to expect

- **Fill colour checks** are wrapped in try/catch — Excel's API can throw
  when reading `.format.fill.color` on an unfilled range instead of
  returning something falsy.
- **Non-adjacent bold checks** run several small `Excel.run` calls in
  parallel (one per cell/range) rather than one batched call — more
  network round-trips, but each check is isolated so one bad property read
  can't take down the others.
- **Live checking is now on by default.** The pane registers `onChanged`
  (fires for both value AND formatting edits — confirmed against current
  Microsoft docs, this isn't a guess), `charts.onAdded`, and
  `worksheets.onAdded` on load, debounced 700ms so a burst of edits doesn't
  trigger a check per keystroke. "Check now" still exists for an immediate
  manual pass (handy right after "Load starter data"). A 6-second poll
  runs alongside the event handlers as a safety net — it's what catches
  the "add a new worksheet" task's data if a student edits a *different*
  sheet than the one active when the handlers were registered, since
  `onChanged` is scoped to one worksheet.
- **Handler cleanup happens on every mode/lesson switch** — verified
  against Microsoft's documented pattern (remove handlers using the same
  `RequestContext` they were added with, not a fresh one). If that removal
  silently fails for some reason, it doesn't corrupt anything visible: a
  render-token guard means any check result from a stale handler gets
  discarded before it touches the DOM, whether or not the underlying Excel
  handler ever actually gets torn down.
- **The dot in the live-status bar** turns green once handlers register
  successfully; if it stays grey, live watching failed silently for that
  session (older Office build, etc.) and it's fallen back to poll-only —
  "Check now" and the 6s poll still work regardless.
- The **new-worksheet check** in the assignment (`/term\s*4/i` name match)
  assumes students name the sheet something containing "Term 4" — if you
  want stricter or looser matching, that regex is the one line to change.