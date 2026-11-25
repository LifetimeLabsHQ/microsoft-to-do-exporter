<p align="center">
  <a href="https://todo.privacynotes.app">
    <img src="public/microsoft-to-do.svg" alt="Microsoft To Do Exporter" width="44" height="44" />
  </a>
</p>

<h1 align="center">Microsoft To Do Exporter</h1>

<p align="center">
  Free, browser-only tool to export your Microsoft To Do lists, tasks, subtasks, notes, and attachments as a local ZIP (JSON, Markdown, and CSV).
</p>

<p align="center">
  <a href="https://todo.privacynotes.app"><strong>Use the Exporter</strong></a>
  &nbsp;&middot;&nbsp;
  <a href="https://privacynotes.app/">PrivacyNotes.app</a>
  &nbsp;&middot;&nbsp;
  <a href="https://try.privacynotes.app/">Live Demo</a>
  &nbsp;&middot;&nbsp;
  <a href="https://privacynotes.app/help/import/microsoft-to-do">Import Guide</a>
</p>

---

## Why we built this

Microsoft announced that Microsoft To Do retires on **June 30, 2027**, with no replacement for personal Microsoft accounts and no built-in export button. Most alternatives only offer cloud-to-cloud importers that lock your tasks into another service.

We built this standalone exporter at **[PrivacyNotes](https://privacynotes.app/)** so anyone - whether switching to PrivacyNotes or another app - can download a clean, portable local backup of their Microsoft To Do data in under a minute.

## Privacy by design

- **100% in your browser.** No backend server, no analytics, no cookies. Your browser connects directly to `login.microsoftonline.com` and `graph.microsoft.com`, then builds the ZIP locally on your device.
- **Strictly read-only.** Requests only `Tasks.Read` via OAuth 2.0 PKCE. Nothing in your Microsoft account is ever edited or deleted.
- **Temporary tokens.** Session tokens live only in `sessionStorage` and clear on sign-out or tab close.

## What is inside the ZIP

| File | Contents |
| --- | --- |
| `microsoft-to-do-export.json` | Complete structured backup (lists, tasks, subtasks, priorities, due dates, repeats, categories, notes, linked items, attachments). |
| `markdown/*.md` | One readable Markdown checklist per list (works in Obsidian, Typora, GitHub, or any editor). |
| `tasks.csv` | Spreadsheet-ready CSV of every task with formula-injection protection. |
| `attachments/*` | Downloaded task attachments, when present. |

## Moving to PrivacyNotes

<p>
  <a href="https://privacynotes.app/">
    <img src="public/privacynotes-icon.svg" alt="PrivacyNotes" width="22" height="24" align="absmiddle" />
  </a>
  &nbsp;<strong><a href="https://privacynotes.app/">PrivacyNotes</a></strong> is an open-source app for notes, tasks, calendar, journal, bookmarks, and vault - where your notes are encrypted on your device, across Web, Windows, Mac, Linux, iOS, and Android.
</p>

Drop the downloaded ZIP into **Settings > Transfer > Import > Microsoft To Do** (or test it right now with no account at [try.privacynotes.app](https://try.privacynotes.app/)). Full walkthrough: [privacynotes.app/help/import/microsoft-to-do](https://privacynotes.app/help/import/microsoft-to-do).

## Run locally

```bash
pnpm install
pnpm dev
```

Open `http://localhost:5175` (add `?demo=1` to preview with sample data without signing in). To use your own Microsoft Entra ID SPA app (`Tasks.Read` delegated permission), set `VITE_MS_CLIENT_ID` in `.env.local`.

```bash
pnpm test
pnpm build
```

## License

MIT
