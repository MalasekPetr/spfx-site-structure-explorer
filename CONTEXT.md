# Blueprint: Site Structure Explorer (SPFx + React + KendoReact)

Build blueprint for the Site Structure Explorer web part. Point Claude Code (or a developer)
at this file and implement strictly to the contract below; ask before deviating. Output ASCII
only. This blueprint accompanies the AI Skills Fest 2026 talk and is referenced by the README
and TUTORIAL - it is the artifact that lets an AI agent scaffold the project fast and stay
inside the guardrails.

Scope note: this is a standalone, public, single-purpose demo. Build toolchain is Heft - SPFx
1.22 replaced the gulp toolchain with Heft (webpack still underneath), so the standard
Heft-only / no-gulp convention applies here; do NOT introduce gulp. Product-specific rules meant
for other repositories (e.g. a license-cache layer) do NOT apply. This file is the authority;
where it conflicts with global conventions, this file wins.

---

## 1. Purpose and framing

Build one SPFx client-side web part that renders a **governance lens** over a single
SharePoint site: every list and library (including hidden and system ones, behind a
toggle), with item counts, nested folders on demand, and a unique-permissions indicator.
The point of the demo is "SPFx + React give a new, useful view over standard SharePoint
data that already exists" - no data seeding, no external store.

Primary visual: **KendoReact TreeList** (self-referencing hierarchical grid). Reliable,
typed, sortable/filterable, matches the governance-table framing.

## 2. Scope

In scope (MVP, must be live-demo safe):
- Enumerate lists + libraries of the current/target site in ONE batched query.
- Columns: Title (+ kind icon), Item count, Hidden, System, Unique permissions, Template.
- Toggle "Show hidden and system" (default OFF).
- Sort + filter via TreeList built-ins.

In scope (stretch, demo only if stable; otherwise pre-recorded):
- Lazy folder drill-down: expand a library/folder -> fetch only that node's child folders.
- Permission detail panel: select a node -> on-demand fetch of role assignments for THAT
  one object, shown in a KendoReact Dialog.

## 3. Non-goals (hard guardrails - do NOT do these)

- NO site-wide permission crawl. Permissions are read per selected object, on demand only.
- NO per-item enumeration to count items. Use the `ItemCount` property on the list/folder.
- NO upfront full folder recursion. Folders load lazily, one expanded node at a time.
- NO Microsoft Graph. Use SPO REST via PnPjs on the delegated SPFx context only.
  (SPO REST via PnPjs is enough and keeps reads in the current user's context.)
- NO product license gate or license-cache layer - this is a standalone, read-only demo,
  not a licensed product.
- NO `any`, NO `@ts-ignore`, NO non-null `!` to silence the compiler.

## 4. Stack and constraints

- Target: SPFx 1.22 / React 17 / TypeScript strict / Node 22. These are the intended pins, but
  DO NOT assume the local toolchain matches - verify (see VERSION DRIFT below).
- Node: use Node 22 (with fnm: `fnm use 22`). Confirm `node --version` is 22.x before scaffolding.
- Build: Heft. SPFx 1.22 replaced the gulp toolchain with Heft (RushStack; webpack still
  underneath) - there is no gulp. Commands run via `heft` / npm scripts:
  `heft start` (serve + hosted workbench), `heft build` (build AND bundle combined),
  `heft build --production` then `heft package-solution --production` (ship + `.sppkg`),
  `heft clean`. Note: `--ship` is gone (now `--production`); there is no separate `bundle` step.
- VERSION DRIFT (resolve before writing any logic): the installed
  `@microsoft/generator-sharepoint` may not be 1.22 (e.g. 1.23.x). The pin that actually matters
  is the REACT version the generator emits, not the SPFx number. After scaffolding (section 12
  step 1), read the generated `app/package.json` and record the real SPFx + React versions. If
  React is not 17, note the deviation as an assumption (section 13); if it cannot be reconciled
  with KendoReact (below), STOP and surface the conflict to the user - never force a React
  mismatch across the build.
- Data: PnPjs 4.x, selective imports only.
- UI: KendoReact (licensed). Packages: `@progress/kendo-react-treelist`,
  `@progress/kendo-react-dialog`, `@progress/kendo-react-inputs` (Switch),
  `@progress/kendo-react-indicators` or `@progress/kendo-react-common` for badges as needed,
  plus a theme package, e.g. `@progress/kendo-theme-default`. Use one theme consistently.
- KendoReact MUST match React: install a KendoReact version whose `peerDependencies` include the
  project's ACTUAL React version. Check the peer range BEFORE `npm install`; do not assume the
  latest KendoReact supports React 17. If no compatible Kendo line exists for the emitted React,
  stop and ask (per VERSION DRIFT).
- KendoReact license: commercial library; activate your own license (trial or paid) via
  `@progress/kendo-licensing` (license file or env var). Do NOT commit the key. See the README.
- Keep utilities small and local; this demo has no external shared-library dependency.
- Single-copy module resolution for `@pnp/*` + `@microsoft/*` is MANDATORY - see section 10.

## 5. Seed project (recommended starting point)

Start from a HYBRID of two stages in `learn-spfx-growth`
(https://github.com/MalasekPetr/learn-spfx-growth):

- Technical base: `4-Tree/app` (web part `assetDeployment`). It already matches this
  blueprint's data layer - SPFx 1.22.2, React 17, Heft, strict TS, and `@pnp/sp ^4.9.0`
  (+ `@pnp/core` / `@pnp/queryable` / `@pnp/logging`) with the `spfi().using(SPFx())` pattern,
  selective imports, OData filtering, and the `services/` + `hooks/` + `models/` layout. It
  also already uses Dialogs/Panels (reuse for the permission detail dialog).
- Reduce to the read-only shape of `3-Plant` (web part `phonelist`). 3-Plant is read-only
  (service + filter + display, no CRUD) - exactly this demo. So after copying 4-Tree:
  - Remove all CRUD (create/update/delete) code paths and lookup-column handling.
  - Remove the multi-table cache and the Dexie dependency. A single `sp.web.lists` call does
    not need caching; keep the data flow deterministic for the live demo.
  - Keep the PnPjs setup, the `services/` + `hooks/` + `models/` structure, and the
    Dialog/Panel pattern.

Do NOT seed from 3-Plant directly: it fetches data via Microsoft Graph (`@microsoft/sp-http`),
which this blueprint forbids (section 3). Seeding from it would mean ripping out Graph and
re-adding PnPjs as step zero.

UI library swap (applies regardless): both stages use Fluent UI. Replace Fluent UI with
KendoReact for the new component (TreeList, Dialog, Switch). Do not keep both UI libraries in
the new web part.

Net path: copy `4-Tree/app` -> rename web part to `siteStructureExplorer` -> strip
CRUD/lookup/cache down to read-only -> swap Fluent UI for KendoReact -> implement per the
sections below.

## 6. Architecture / layout

```
src/webparts/siteStructureExplorer/
  SiteStructureExplorerWebPart.ts        // SPFx web part, props, passes sp + context
  components/
    SiteStructureExplorer.tsx            // top-level component, state, toggle, dialog
    StructureTreeList.tsx                // KendoReact TreeList wiring (columns, expand)
    PermissionDetailDialog.tsx           // on-demand role-assignment panel
    cells/                               // small cell renderers (BooleanBadge, KindIcon)
  models/
    ISiteStructureNode.ts
    IPermissionAssignment.ts
    enums.ts                             // NodeKind, PrincipalTypeMap, BaseTemplate labels
  services/
    ISiteStructureService.ts             // interface (the contract)
    SiteStructureService.ts              // PnPjs implementation
```

Layering: WebPart -> Component (state) -> Service (interface) -> PnPjs. UI never calls
PnPjs directly; it goes through `ISiteStructureService`. The service returns typed models
only - no raw PnPjs shapes leak into the UI.

## 7. Data model (write these types FIRST - they are the guardrails)

```ts
// enums.ts
export type NodeKind = "site" | "list" | "library" | "folder";

// Primary system signal is `Hidden || IsCatalog` (catches almost everything, incl. all
// *Catalog templates). This set is the supplementary catch for infrastructure lists that
// may be visible and non-catalog. Values are SPListTemplateType (BaseTemplate) IDs,
// verified against Microsoft Learn (SPListTemplateType enumeration). Tune to taste.
export const KNOWN_SYSTEM_TEMPLATES: ReadonlySet<number> = new Set<number>([
  110, // DataSources
  111, // WebTemplateCatalog (site template gallery)
  112, // UserInformation (User Information List)
  113, // WebPartCatalog (Web Part gallery)
  114, // ListTemplateCatalog (List Template gallery)
  116, // MasterPageCatalog (Master Page gallery)
  117, // NoCodeWorkflows
  118, // WorkflowProcess
  121, // SolutionCatalog (Solutions)
  122, // NoCodePublic (No Code Public Workflow)
  123, // ThemeCatalog (Themes)
  124, // DesignCatalog
  125, // AppDataCatalog
  140, // WorkflowHistory
  151, // HelpLibrary
  160, // AccessRequest (Access Requests list)
  175, // MaintenanceLogs
  1200, // AdminTasks
  1220, // HealthRules
  1221, // HealthReports
  1230, // DeveloperSiteDraftApps
]);
// NOTE: deliberately NOT system (these are user content): 100 GenericList, 101 DocumentLibrary,
// 102-109 standard lists, 115 XMLForm, 119 WebPageLibrary (Site Pages), 120 CustomGrid,
// 130 DataConnectionLibrary, 600 ExternalList, 1100 IssueTracking, 700 MySiteDocumentLibrary.
// For the Template column label, derive a friendly name from the same BaseTemplate id
// (small `BASE_TEMPLATE_LABELS: Record<number,string>` map; folders show "-"). Unknown ids
// fall back to "Template <id>".

// ISiteStructureNode.ts
export interface ISiteStructureNode {
  id: string;                       // stable key: list GUID, or folder server-relative URL
  parentId: string | null;         // self-referencing key for TreeList; site root = null
  kind: NodeKind;
  title: string;
  serverRelativeUrl: string;
  itemCount: number;               // list.ItemCount or folder.ItemCount (cheap, no enumeration)
  hidden: boolean;                 // list.Hidden; folders inherit parent's flag for display
  isSystem: boolean;               // derived (see rule in section 8)
  baseTemplate: number;            // list.BaseTemplate (folders: -1)
  hasUniqueRoleAssignments: boolean;
  hasChildren: boolean;            // library or folder that may contain subfolders -> lazy expand
  loaded: boolean;                 // lazy-load state for TreeList expand
  expanded?: boolean;
}

// IPermissionAssignment.ts
export type PrincipalKind = "User" | "SharePointGroup" | "SecurityGroup" | "Unknown";

export interface IPermissionAssignment {
  principalName: string;
  principalKind: PrincipalKind;
  roles: string[];                 // RoleDefinitionBindings names, "Limited Access" excluded
}
```

## 8. Service contract and queries

```ts
// ISiteStructureService.ts
export interface ISiteStructureService {
  getSiteRoot(): Promise<ISiteStructureNode>;
  getListsAndLibraries(opts: { includeHiddenSystem: boolean }): Promise<ISiteStructureNode[]>;
  getChildFolders(node: ISiteStructureNode): Promise<ISiteStructureNode[]>; // lazy, one node
  getPermissions(node: ISiteStructureNode): Promise<IPermissionAssignment[]>; // on-demand
}
```

PnPjs setup (selective imports):
```ts
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/folders";
import "@pnp/sp/security";
// const sp = spfi().using(SPFx(this.context));  // built in the web part, injected into service
```

`getListsAndLibraries` - ONE batched call, no per-item work:
```ts
sp.web.lists
  .select(
    "Id", "Title", "Hidden", "IsCatalog", "BaseTemplate", "BaseType",
    "ItemCount", "HasUniqueRoleAssignments",
    "RootFolder/ServerRelativeUrl"
  )
  .expand("RootFolder")();
```
- Map each list to an `ISiteStructureNode` with `parentId = <siteRoot id>`.
- `kind`: BaseType 1 -> "library", else "list".
- `isSystem = hidden || isCatalog || KNOWN_SYSTEM_TEMPLATES.has(baseTemplate)`.
  Keep a small `KNOWN_SYSTEM_TEMPLATES` set (e.g. user info list, catalogs, form templates);
  document the chosen set in code comments.
- `hasChildren = kind === "library"` (libraries can hold folders; offer lazy expand).
- When `opts.includeHiddenSystem` is false, filter out `isSystem` nodes before returning.

`getSiteRoot`:
```ts
sp.web.select("Id", "Title", "ServerRelativeUrl", "HasUniqueRoleAssignments")();
// -> single root node, parentId null, kind "site"
```

`getChildFolders` (lazy, only for the expanded node):
```ts
sp.web.getFolderByServerRelativePath(node.serverRelativeUrl).folders
  .select("Name", "ServerRelativeUrl", "ItemCount", "TimeLastModified")();
```
- Map each subfolder to a node: `parentId = node.id`, `kind = "folder"`, `baseTemplate = -1`,
  `hasUniqueRoleAssignments` left false unless cheaply available (do not expand here).
- Filter the library system folder "Forms" (and similar) or flag it as system.
- `hasChildren` for a folder: set true if it itself reports subfolders; simplest is to
  attempt-on-expand and set `loaded` after fetch. Keep it lazy - never pre-walk.

`getPermissions` (on demand, single object):
```ts
sp.web.lists.getById(node.id).roleAssignments
  .expand("Member", "RoleDefinitionBindings")();
```
- Map `Member.PrincipalType` (1 User, 4 SecurityGroup, 8 SharePointGroup) to `PrincipalKind`.
- `roles` = `RoleDefinitionBindings[].Name`, excluding "Limited Access".
- IMPORTANT caveat to surface in the UI: results reflect only what the CURRENT USER is
  allowed to see. Render a short note in the dialog stating this.

Throttling: keep calls minimal (one list call; folders/permissions only on interaction).
Enable PnPjs retry behavior on the `spfi` instance. On any 429/error, show a graceful
inline error state, never a blank component.

## 9. UI behavior (KendoReact TreeList)

- Bind flat `ISiteStructureNode[]` using TreeList self-referencing data (id / parentId).
  Use the TreeList flat-data-to-tree helper; `subItemsField` virtualized via expand.
- Columns:
  - Title: custom cell with `KindIcon` (site/list/library/folder) + text.
  - Item count: numeric, right-aligned, sortable.
  - Hidden: `BooleanBadge`.
  - System: `BooleanBadge`.
  - Unique permissions: `BooleanBadge`; clicking it (or selecting the row) opens
    `PermissionDetailDialog` for that node.
  - Template: friendly label from BaseTemplate map (folders show "-").
- Toolbar: KendoReact `Switch` "Show hidden and system" -> re-filters the dataset
  (re-run `getListsAndLibraries({ includeHiddenSystem })`, or filter client-side from a
  full fetch held in state; prefer a single full fetch + client filter to avoid refetch).
- Expand: `onExpandChange` -> if node.kind is library/folder and not `loaded`, call
  `getChildFolders(node)`, append children to state, mark `loaded = true`.
- Selection -> `PermissionDetailDialog`: calls `getPermissions(node)` on open; shows a
  loading spinner, then the principal/role table, plus the visibility caveat note.
- States: loading (skeleton or spinner), empty (no lists matched the filter), error (inline).

## 10. Build / config specifics

- Toolchain: Heft (SPFx 1.22). The 1.22 generator scaffolds a Heft project by default - no gulp.
- MANDATORY single-copy resolution for `@pnp/*` (the #1 trap). PnPjs v4 selective-import
  augmentation breaks if more than one copy of `@pnp/sp` is resolved: `.lists`, `.folders`,
  `.roleAssignments` come back undefined or "property does not exist on type". Fix it by
  REPLACING the generated `tsconfig.json` with one that extends the Heft web-build-rig base and
  pins `@pnp/*` to the app's single copy via `paths`:

  ```jsonc
  {
    "extends": "./node_modules/@microsoft/spfx-web-build-rig/profiles/default/tsconfig-base.json",
    "compilerOptions": {
      "target": "ES2022",
      "lib": ["ES2022", "DOM"],
      "module": "ESNext",
      "moduleResolution": "node",
      "jsx": "react",
      "declaration": true,
      "sourceMap": true,
      "experimentalDecorators": true,
      "strictNullChecks": true,
      "skipLibCheck": true,
      "outDir": "lib",
      "noImplicitAny": true,
      "forceConsistentCasingInFileNames": true,
      "resolveJsonModule": true,
      "baseUrl": ".",
      "paths": {
        "@pnp/sp": ["./node_modules/@pnp/sp"],
        "@pnp/sp/*": ["./node_modules/@pnp/sp/*"],
        "@pnp/core": ["./node_modules/@pnp/core"],
        "@pnp/core/*": ["./node_modules/@pnp/core/*"]
      }
    },
    "include": ["src/**/*.ts", "src/**/*.tsx"]
  }
  ```

  If duplicate `@pnp/*` copies still appear after install, add a package.json `overrides` block
  pinning `@pnp/*` to one version. Verify a trivial `sp.web()` plus one augmented call
  (`sp.web.lists`) type-checks and resolves before building the service.
- KendoReact theme: import one theme (e.g. `@progress/kendo-theme-default`); do not mix themes.
- KendoReact license activation wired into the build (env/license file); never commit secrets.
- Web part property: target site is the current site by default; optionally expose a
  "site URL" text property for pointing at another site in the same tenant (read-only access).

## 11. Definition of done (acceptance criteria)

Two buckets. The agent must get the first bucket green, and must NOT claim the second as done -
it has no tenant auth or license key headless. It hands those off to the user explicitly.

Agent-verifiable (all must be green before handoff):
1. `heft build` is clean; `heft build --production` then `heft package-solution --production`
   are clean (Heft combines build + bundle; there is no `gulp` and no separate bundle step).
2. eslint clean; strict TS clean - no `any`, no `@ts-ignore`, no `!` used to silence the compiler.
3. Lists + libraries load via a SINGLE `sp.web.lists` call; counts come from `ItemCount`
   (no per-item enumeration anywhere in the code).
4. No Microsoft Graph anywhere; SPO REST via PnPjs on the delegated SPFx context only.
5. Folder expand lazy-loads only the expanded node; no upfront recursion exists in code.
6. Unique-permissions is a boolean column; permission detail is fetched only on demand for the
   selected node; the current-user-visibility caveat is rendered in the dialog.
7. "Show hidden and system" toggle works; default hides them.
8. `KNOWN_SYSTEM_TEMPLATES` present; review confirms overrides / tsconfig pins, no dead code.
9. Throttling-safe: PnPjs retry enabled; graceful error and empty states; never a blank UI.

Human-verifiable (agent hands off, states plainly, does NOT mark done):
- Renders in the hosted workbench against a real site (needs the user's tenant auth). To run
  locally: `heft trust-dev-cert` once, then `heft start`, then open the hosted workbench.
- KendoReact renders without a trial banner (needs an activated license key).
- Permission detail reflects what the signed-in user is allowed to see.

## 12. Implementation order for Claude Code

0. Confirm Node 22 (`node --version`; `fnm use 22` if needed).
1. Scaffold into `app/`, non-interactively and WITHOUT install yet. Run from inside `app/`:
   `yo @microsoft/sharepoint --skip-install --solution-name site-structure-explorer
   --component-type webpart --component-name SiteStructureExplorer
   --component-description "Governance lens over a SharePoint site" --framework react
   --environment spo`
   If the command stalls, it is waiting on an UNFLAGGED stdin prompt (e.g. tenant-wide deploy) -
   answer it; never leave it hanging.
   The 4-Tree seed (section 5) is OPTIONAL; a fresh scaffold is fine and often cleaner here. If
   you skip the seed, you OWN the section 10 single-copy setup yourself. Record the seed-skip as
   an assumption (section 13).
2. Read `app/package.json`: record actual SPFx + React versions. Choose a KendoReact version
   whose peer range includes that React (section 4). If React is not 17, note it or stop per the
   VERSION DRIFT rule.
3. Replace `tsconfig.json` per section 10 (extends the Heft rig base + `@pnp/*` single-copy
   `paths`) FIRST. Then add the KendoReact packages (React-matched) + one theme + license activation.
4. `npm install` once. Run `heft build` (`npm run build`) and confirm the EMPTY shell compiles
   before writing any logic. This gate is mandatory - do not proceed on a non-compiling shell.
5. Write the typed models (section 7) - the contract first.
6. Implement `getSiteRoot` + `getListsAndLibraries`; bind a plain TreeList. This is the
   live-safe MVP.
7. Add the "Show hidden and system" Switch + client-side filtering.
8. Add lazy folder expand (`getChildFolders` + `onExpandChange`).
9. Add `PermissionDetailDialog` (`getPermissions` on open) with the visibility caveat.
10. Polish loading/empty/error states. Get eslint + `heft build` green.

Commit after each step. Keep each step independently runnable so the demo can stop at the
MVP (step 6). Hand off the live-workbench and license checks to the user (section 11).

## 13. Decided parameters (CC: do NOT ask - use these)

CC must not pause to ask the presenter about anything in this section, and must not block on
interactive clarification. Use the stated default, proceed, and note the assumption in the
commit message. For any ambiguity NOT covered here, pick the most conservative option
consistent with sections 1-3 (read-only, no Graph, no crawl) and record it in the commit.

- Target site URL: DEFAULT = current site (`this.context.pageContext.web.absoluteUrl`).
  Expose an optional web-part text property "Site URL"; when empty, use the current site.
  Do not ask - just default to current site.
- KendoReact theme: do NOT ask. Default to `@progress/kendo-theme-default` (or, if you are
  seeding into a workspace that already standardizes on a theme, reuse that one).
- Hero visual (OrgChart): NOT in scope. Do not implement it and do not ask about it. The
  TreeList is the only visualization for this build.
- "Limited Access" role: excluded from the permission display. Fixed - not a question.

## 14. Demo notes (not for CC - for the presenter)

- Live-safe path: steps 3-4 (one call, toggle). Deterministic regardless of site size.
- Stretch live (only if stable): one folder expand + one permission dialog.
- Pre-record: the Claude Code scaffolding session itself (the velocity moment) and any deep
  folder recursion. Reuse the timebox rule: if a live fetch stalls, cut to the recording.
