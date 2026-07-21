# Stream-Flo CSAT Dashboard (SPFx)

Live Field Service customer-satisfaction dashboard, delivered as a SharePoint
Framework (SPFx) client-side web part. It reads the **CSAT RESPONSES** list
**directly** via SharePoint REST under the signed-in user's O365 context — no
Power Automate read flow, no anonymous URL, no stored credentials. Access to
the dashboard is whatever access the user already has to the list/site.

This is the authenticated, internal counterpart to the public survey (which is
an anonymous static site posting through an intake flow). Clean split: public
intake, private analytics.

## What it shows

KPIs (total responses, avg CSAT, NPS, safety) with period-over-period deltas,
a "needs attention" list (low CSAT / low safety / detractor NPS), rating
breakdown by category, NPS breakdown, CSAT trend line, per-branch bar chart,
most-cited improvement areas, technician performance table, and verbatim
exemplary / improvement comments. Branch / date-range / technician filters and
CSV + Excel export are built in.

## Tech

- SPFx **1.18.2**, no-framework web part, TypeScript 4.7, Node 18.
- Chart.js 4.4.1 and SheetJS (xlsx) 0.18.5 loaded from CDN via
  `SPComponentLoader`. If a tenant blocks the CDN, the dashboard still renders;
  charts / Excel export degrade gracefully (CSV export is dependency-free).
- Data access: `src/webparts/csatDashboard/CsatDataService.ts`
  (`SPHttpClient`, follows `@odata.nextLink` paging, maps the list's `serviceNptes`
  typo column back to `serviceNotes`).

## Web part properties

Configure in the property pane after adding the web part:

- **List site URL** — full URL of the site holding the list. Defaults to
  `https://streamflogroup.sharepoint.com/sites/SUSTeam-SUSBranchOperations`.
  Leave blank to use the current site. Cross-site reads work as long as the
  viewer has permission to that list.
- **List name** — defaults to `CSAT RESPONSES`.
- **Auto-refresh interval** — seconds; 0 = off.

## Building the package

### Option A — GitHub Actions (recommended)

Every push to `main` runs `.github/workflows/build.yml` on Node 18, which
bundles and packages the solution and uploads the built **`.sppkg`** as a
workflow artifact. Grab it from the run's *Artifacts* section — no local build
needed (handy since local Node here is v24, but SPFx 1.18 needs Node 18).

You can also trigger it manually from the Actions tab (**Run workflow**).

### Option B — local build (needs Node 18)

```bash
npm install
gulp bundle --ship
gulp package-solution --ship
# -> sharepoint/solution/spfx-csat-dashboard.sppkg
```

## Deploying

1. Upload `spfx-csat-dashboard.sppkg` to your **tenant App Catalog**
   (`/sites/appcatalog` → *Apps for SharePoint*).
2. When prompted, trust the solution. It's tenant-scoped and needs no special
   API permissions — it reads the list with the signed-in user's own rights
   via SPHttpClient.
3. On the SUS Branch Operations site (or wherever you want it), **Add an app**
   → *spfx-csat-dashboard*.
4. Edit a modern page → add the **Stream-Flo CSAT Dashboard** web part (under
   the *Advanced* group) → set the list site URL / name in the property pane if
   they differ from the defaults → publish.

## Notes

- Because the web part uses the current user's context, the dashboard honors
  SharePoint permissions automatically — only people who can see the list can
  see the data.
- The old "CSAT Dashboard — Read Responses" Power Automate flow is no longer
  needed by this dashboard and can be retired once you've cut over.
