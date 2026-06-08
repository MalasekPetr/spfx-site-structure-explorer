# Site Structure Explorer

An SPFx + React + KendoReact web part that turns a SharePoint site's own metadata into an
interactive **governance lens**: every list and library (including hidden and system ones,
behind a toggle), with item counts, folders on demand, and a unique-permissions flag.

No data seeding, no external store - it only reads metadata that already exists in the site.
Built as a conference demo for **AI Skills Fest 2026** ("Vyvojar + AI = 10x?") to show two
things: SPFx + React give genuinely new, useful views over standard SharePoint data, and a
tight blueprint lets an AI agent scaffold that view fast.

> Status: demo / teaching project. Read-only. Not a production product.

## Features

- One interactive grid of all lists and libraries on a site (KendoReact TreeList).
- Columns: Title (+ kind icon), Item count, Hidden, System, Unique permissions, Template.
- Toggle to show or hide system and hidden lists (off by default).
- Lazy folder drill-down - expand a library to load only that node's child folders.
- On-demand permission detail - select a row to see role assignments for that one object.
- Sorting and filtering via the TreeList built-ins.

## Why it is interesting

The data is the least glamorous part of SharePoint - it is just the site's own structure.
The point is that a thin SPFx + React layer turns it into a view you cannot get out of the
box: where the data actually lives, what is hidden, and what has broken permission
inheritance. Same standard data, new lens.

## Tech stack

- SharePoint Framework (SPFx) 1.22, web part, React 17, TypeScript (strict).
- [PnPjs](https://pnp.github.io/pnpjs/) 4.x for SharePoint REST (delegated SPFx context).
- [KendoReact](https://www.telerik.com/kendo-react-ui) for the UI (TreeList, Dialog, Switch).
- No Microsoft Graph: all data comes from SPO REST via PnPjs on the current user's context.

## Design notes and honest limitations

- **Item counts are cheap.** They come from the `ItemCount` property on the list / folder
  object - the code never enumerates items to count them.
- **Folders load lazily.** There is no upfront recursion across the whole site; only the
  node you expand is fetched. This keeps it throttling-safe and fast on large sites.
- **Permissions are read on demand, per object.** The grid shows the cheap
  `HasUniqueRoleAssignments` boolean; full role assignments are fetched only for the row you
  open. There is no site-wide permission crawl.
- **Permission visibility is scoped to you.** Role-assignment results only reflect what the
  current user is allowed to see. On a site where you are not an admin, the detail may be
  partial. The dialog states this.
- **System detection.** The primary signal is `Hidden || IsCatalog`; a curated
  `KNOWN_SYSTEM_TEMPLATES` set (BaseTemplate IDs) catches the rest. See `src/.../models/enums.ts`.

## KendoReact license (read before you build)

KendoReact is a **commercial** library. The npm packages are public, but using them without a
license shows a trial banner / requires an active license key.

- This repository is MIT licensed and contains **no** license key.
- To build it you need your own KendoReact license: a free 30-day trial or a paid/DevCraft
  license. See the official
  [KendoReact licensing docs](https://www.telerik.com/kendo-react-ui/components/my-license/).
- The MIT license applies to the source in this repo only, not to KendoReact.

## Prerequisites

- Node.js 22 LTS (matches the SPFx 1.22 compatibility matrix).
- A Microsoft 365 / SharePoint Online tenant where you can use the hosted workbench and,
  for deployment, the App Catalog.
- A KendoReact license (trial is fine).

## Quick start

```bash
git clone https://github.com/MalasekPetr/spfx-site-structure-explorer.git
cd spfx-site-structure-explorer/app
npm install

# Activate your KendoReact license (see Telerik docs; do NOT commit the key)
# e.g. place your license file in the project root, then run the activation script.

gulp trust-dev-cert        # first time only
gulp serve --nobrowser
```

Then open the hosted workbench and add the web part:

```
https://<your-tenant>.sharepoint.com/_layouts/15/workbench.aspx
```

By default the web part reads the **current** site. An optional "Site URL" property lets you
point it at another site in the same tenant (read access required).

## Deploy

```bash
gulp bundle --ship
gulp package-solution --ship
# upload sharepoint/solution/*.sppkg to your tenant App Catalog
```

## How it is built

The UI never calls PnPjs directly - it goes through a typed `ISiteStructureService`, which
returns typed models only. Layering: `WebPart -> Component (state) -> Service -> PnPjs`.
Types are written first and act as guardrails for both humans and the AI agent.

This project was scaffolded with **Claude Code from a written blueprint**
([`CONTEXT.md`](./CONTEXT.md)) - the same "blueprint first, types as guardrails" approach the
talk argues makes AI assistance actually fast. To reproduce the whole thing from scratch with
only publicly available components, follow [`TUTORIAL.md`](./TUTORIAL.md).

## Reproduce from scratch

See **[TUTORIAL.md](./TUTORIAL.md)** for the full step-by-step (scaffold, PnPjs, KendoReact,
models, service, TreeList, lazy folders, permission dialog, deploy).

## License

MIT (c) Petr Malasek / Malach IS. See [LICENSE](./LICENSE).
KendoReact is licensed separately by Progress / Telerik and is not covered by this license.
