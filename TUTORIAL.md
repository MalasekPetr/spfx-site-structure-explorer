# TUTORIAL: Build Site Structure Explorer from scratch

This reproduces the whole web part using only publicly available components. Every tool here
is public; KendoReact additionally needs your own license key (a free 30-day trial works).

Target stack: SPFx 1.22, React 17, TypeScript strict, PnPjs 4.x, KendoReact.

> Two paths: the manual path below (no AI), or the fast path - point Claude Code at the
> blueprint (`CONTEXT.md`) and let it scaffold. The manual path is the reference.

---

## 0. Prerequisites

- Node.js 22 LTS. Check the SPFx 1.22 compatibility matrix if unsure.
- A SharePoint Online tenant (hosted workbench for testing; App Catalog for deployment).
- A KendoReact license (trial or paid): https://www.telerik.com/kendo-react-ui/components/my-license/

Install the global toolchain (SPFx 1.22 uses the Heft toolchain - no gulp-cli needed; Heft comes
as a local dependency in the scaffolded project):

```bash
npm install -g yo @microsoft/generator-sharepoint
```

---

## 1. Scaffold the SPFx web part

Scaffold non-interactively, create a **webpart** component, and **skip the install** for now -
dependencies and the tsconfig are set in steps 2-3, then installed once at the end of step 3:

```bash
mkdir spfx-site-structure-explorer && cd spfx-site-structure-explorer
mkdir app && cd app
yo @microsoft/sharepoint --skip-install \
  --solution-name site-structure-explorer \
  --component-type webpart \
  --component-name SiteStructureExplorer \
  --component-description "Governance lens over a SharePoint site" \
  --framework react \
  --environment spo
```

If the command stalls, it is waiting on an unflagged prompt (e.g. tenant-wide deploy) - answer
it; do not leave it hanging. When it finishes, open `app/package.json` and note the actual SPFx
and React versions. The **React** version is the one that matters: KendoReact must match it
(step 3). SPFx 1.22 scaffolds a Heft project (no gulp).

---

## 2. Add PnPjs and wire it to the SPFx context

Add these PnPjs packages (installed once at the end of step 3, not now):

```
@pnp/sp  @pnp/core  @pnp/queryable  @pnp/logging
```

In the web part class, create one `sp` instance from the SPFx context and pass it into the
component (selective imports keep the bundle small):

```ts
// SiteStructureExplorerWebPart.ts
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/folders";
import "@pnp/sp/security";

private _sp: SPFI;

protected onInit(): Promise<void> {
  this._sp = spfi().using(SPFx(this.context));
  return super.onInit();
}
// pass this._sp (and the optional siteUrl property) into the React component via props
```

---

## 3. Add KendoReact + theme + license

```bash
npm install \
  @progress/kendo-react-treelist \
  @progress/kendo-react-dialog \
  @progress/kendo-react-inputs \
  @progress/kendo-react-buttons \
  @progress/kendo-react-common \
  @progress/kendo-data-query \
  @progress/kendo-licensing \
  @progress/kendo-theme-default
```

Activate your license (do NOT commit the key). The current flow: place your license file in
the project root and run the activation step from the Telerik docs linked above. The key can
also be supplied via an environment variable in CI.

Import the theme once (e.g. at the top of the web part or root component):

```ts
import "@progress/kendo-theme-default/dist/all.css";
```

Gotchas: SPFx ships its own React (17 for 1.22) - do not add a second React. Install a
KendoReact version whose `peerDependencies` include that React (do not assume the latest Kendo
supports React 17). If the theme CSS does not load, confirm SPFx is bundling the imported
`.css` (import it from a `.ts/.tsx`, not only SCSS).

### Single-copy tsconfig, then install once and build the empty shell

Before the first PnPjs build, replace the generated `tsconfig.json` so `@pnp/*` resolves to a
single copy (otherwise PnPjs v4 augmentation breaks - `.lists` / `.folders` / `.roleAssignments`
come back undefined):

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

Now install everything once and confirm the empty shell builds before writing any logic:

```bash
npm install
heft build              # build + bundle (combined in Heft); must be clean
heft trust-dev-cert     # first machine only
heft start              # serve; open /_layouts/15/workbench.aspx and add the web part
```

(`heft` runs via the project's npm scripts - `npm run build`, `npm run start` - or `npx heft <task>`.)

---

## 4. Define the typed models first (the guardrails)

```ts
// models/enums.ts
export type NodeKind = "site" | "list" | "library" | "folder";

// Primary system signal is `Hidden || IsCatalog`. This set is the supplementary catch for
// infrastructure lists that may be visible and non-catalog. SPListTemplateType (BaseTemplate)
// IDs, verified against Microsoft Learn.
export const KNOWN_SYSTEM_TEMPLATES: ReadonlySet<number> = new Set<number>([
  110, 111, 112, 113, 114, 116, 117, 118, 121, 122, 123, 124, 125,
  140, 151, 160, 175, 1200, 1220, 1221, 1230,
]);
// Deliberately NOT system (user content): 100 GenericList, 101 DocumentLibrary, 102-109,
// 115 XMLForm, 119 Site Pages, 120 CustomGrid, 130 DataConnectionLibrary, 600 ExternalList,
// 700 MySiteDocumentLibrary, 1100 IssueTracking.
```

```ts
// models/ISiteStructureNode.ts
export interface ISiteStructureNode {
  id: string;                      // list GUID, or folder server-relative URL
  parentId: string | null;        // self-referencing key; site root = null
  kind: import("./enums").NodeKind;
  title: string;
  serverRelativeUrl: string;
  itemCount: number;               // from ItemCount (no enumeration)
  hidden: boolean;
  isSystem: boolean;               // derived (see service)
  baseTemplate: number;            // folders: -1
  hasUniqueRoleAssignments: boolean;
  hasChildren: boolean;            // library/folder that may contain subfolders -> lazy expand
  loaded: boolean;                 // lazy-load state
  expanded?: boolean;
  children?: ISiteStructureNode[]; // filled lazily for TreeList
}
```

```ts
// models/IPermissionAssignment.ts
export type PrincipalKind = "User" | "SharePointGroup" | "SecurityGroup" | "Unknown";
export interface IPermissionAssignment {
  principalName: string;
  principalKind: PrincipalKind;
  roles: string[];                 // RoleDefinitionBindings names, "Limited Access" excluded
}
```

---

## 5. Implement the service (PnPjs)

```ts
// services/ISiteStructureService.ts
export interface ISiteStructureService {
  getSiteRoot(): Promise<ISiteStructureNode>;
  getListsAndLibraries(opts: { includeHiddenSystem: boolean }): Promise<ISiteStructureNode[]>;
  getChildFolders(node: ISiteStructureNode): Promise<ISiteStructureNode[]>; // lazy
  getPermissions(node: ISiteStructureNode): Promise<IPermissionAssignment[]>; // on demand
}
```

Lists and libraries - ONE batched call, counts from `ItemCount`:

```ts
const raw = await sp.web.lists
  .select("Id","Title","Hidden","IsCatalog","BaseTemplate","BaseType",
          "ItemCount","HasUniqueRoleAssignments","RootFolder/ServerRelativeUrl")
  .expand("RootFolder")();

const nodes = raw.map(l => ({
  id: l.Id,
  parentId: siteRootId,
  kind: l.BaseType === 1 ? "library" : "list",
  title: l.Title,
  serverRelativeUrl: l.RootFolder.ServerRelativeUrl,
  itemCount: l.ItemCount,
  hidden: l.Hidden,
  isSystem: l.Hidden || l.IsCatalog || KNOWN_SYSTEM_TEMPLATES.has(l.BaseTemplate),
  baseTemplate: l.BaseTemplate,
  hasUniqueRoleAssignments: l.HasUniqueRoleAssignments,
  hasChildren: l.BaseType === 1,
  loaded: false,
} as ISiteStructureNode));
// if !opts.includeHiddenSystem -> filter out isSystem before returning
```

Lazy child folders (only for the expanded node):

```ts
const folders = await sp.web
  .getFolderByServerRelativePath(node.serverRelativeUrl).folders
  .select("Name","ServerRelativeUrl","ItemCount","TimeLastModified")();
// map to nodes: parentId = node.id, kind "folder", baseTemplate -1; flag/skip "Forms"
```

Permissions for one object, on demand:

```ts
const ras = await sp.web.lists.getById(node.id).roleAssignments
  .expand("Member","RoleDefinitionBindings")();
// PrincipalType: 1 User, 4 SecurityGroup, 8 SharePointGroup
// roles = RoleDefinitionBindings[].Name, excluding "Limited Access"
```

Enable PnPjs retry on the `spfi` instance and wrap calls so a 429 / error shows a graceful
inline state, never a blank component.

---

## 6. Build the TreeList UI

- Bind the flat nodes as a self-referencing tree (TreeList uses `subItemsField` for children;
  build each node's `children` array from `parentId`, or use the TreeList flat-to-tree helper).
- Columns: Title (custom cell with a kind icon), Item count (numeric, sortable),
  Hidden (boolean badge), System (boolean badge), Unique permissions (boolean badge),
  Template (friendly label from a small BaseTemplate map; folders show "-").
- Toolbar: a KendoReact `Switch` "Show hidden and system". Prefer one full fetch held in
  component state + client-side filter, so toggling does not refetch.

```tsx
<TreeList
  data={treeData}
  columns={columns}
  expandField="expanded"
  subItemsField="children"
  onExpandChange={onExpandChange}
  onRowClick={onRowClick}
/>
```

---

## 7. Lazy folder expand

```ts
const onExpandChange = async (e) => {
  const node = e.dataItem as ISiteStructureNode;
  node.expanded = !node.expanded;
  if (node.expanded && (node.kind === "library" || node.kind === "folder") && !node.loaded) {
    node.children = await service.getChildFolders(node);
    node.loaded = true;
  }
  setTreeData([...treeData]); // re-render
};
```

Never pre-walk the whole site - fetch children only when a node is expanded.

---

## 8. Permission detail dialog

- On row click / clicking the unique-permissions badge, open a KendoReact `Dialog`.
- On open, call `service.getPermissions(node)`; show a spinner, then the principal/role table.
- Render a short note: results reflect only what the current user is allowed to see.

---

## 9. Run, package, deploy

```bash
heft start
# test in the hosted workbench: /_layouts/15/workbench.aspx

heft build --production            # build + bundle for release
heft package-solution --production # produces the .sppkg
# upload sharepoint/solution/*.sppkg to the tenant App Catalog, then add the web part to a page
```

---

## Reproduce checklist

- [ ] Web part renders in the hosted workbench against the current site.
- [ ] Lists/libraries load via a single `sp.web.lists` call; counts from `ItemCount`.
- [ ] Show-hidden-and-system toggle works; default hides them.
- [ ] Folder expand lazy-loads only the expanded node.
- [ ] Unique-permissions boolean shown; detail fetched only on demand; visibility caveat shown.
- [ ] No Microsoft Graph calls; SPO REST via PnPjs only.
- [ ] Strict TS clean: no `any`, no `@ts-ignore`. eslint passes.

---

## Faster path: scaffold with AI from the blueprint

Everything above is captured as a written blueprint in `CONTEXT.md` (purpose, non-goals,
typed contract, service queries, UI behavior, definition of done). Point Claude Code at it and
let it scaffold step by step. The blueprint - not the chat - is what makes the AI run fast and
stay inside the guardrails. That is the whole argument of the talk this repo accompanies.
