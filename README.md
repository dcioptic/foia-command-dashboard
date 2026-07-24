# Division of Criminal Investigation FOIA Command Dashboard

Static dashboard prototype for tracking Illinois State Police FOIA requests.

This dashboard now supports loading data from SharePoint through Microsoft Graph.
If Graph authentication is not configured or fails, it falls back to `foia-data.json`.

## Files

- `index.html` - dashboard layout and UI containers
- `style.css` - executive-style law enforcement dashboard theme
- `app.js` - data loading, KPI calculations, table rendering, and interactions
- `foia-data.json` - sample FOIA request records
- `graph-config.js` - local Graph auth configuration (token or MSAL settings)

## Features

- Header: Division of Criminal Investigation FOIA Command Dashboard
- KPI cards:
	- Total Open FOIAs
	- Due in 10 Days
	- Due in 5 Days
	- Due Today
	- Overdue
	- Average Days Open
	- Closed This Month
	- Open Work Units
- DCI Work Unit filter with "All DCI Work Units"
- DCI Work Unit summary with open and deadline buckets
- Upcoming deadlines panel
- Searchable FOIA request table
- Right-side detail panel for selected request
- Dark mode toggle
- Presentation mode toggle
- SharePoint list resolution by URL path match (`/Lists/DCI%20FOIA/`)
- Full list item retrieval with Graph pagination (`@odata.nextLink`)

## SharePoint / Graph Connection

Target site:

- `https://ilgov.sharepoint.com/teams/ISP.DCI.Commanders`

Target list path:

- `/Lists/DCI%20FOIA/`

Expected display name:

- `DCI FOIA`

### 1) Configure authentication in `graph-config.js`

Choose one method:

- Set `accessToken` directly for quick testing.
- Or set `clientId` and `tenantId` so MSAL can sign in and acquire a Graph token.

Required Graph permission scope:

- `Sites.Read.All`

### 2) Start the dashboard

```bash
python3 -m http.server 8080
```

Open:

- `http://localhost:8080`

### 3) Verify list resolution

Open the browser console. On successful connection, the app logs:

- resolved list `id`
- resolved list `displayName`
- resolved list `webUrl`

The app resolves the site with Graph, enumerates all site lists, and selects the list whose `webUrl` contains `/Lists/DCI%20FOIA/` (preferred), with display name as fallback.

After list resolution, the app retrieves all accessible list items from:

- `/sites/{site-id}/lists/{list-id}/items?$expand=fields`

It follows `@odata.nextLink` until every page is loaded, then applies dashboard filtering locally.

## Run Locally

Because the dashboard fetches JSON, run it from a local web server.

Option 1 (Python):

```bash
python3 -m http.server 8080
```

Then open:

`http://localhost:8080`

## Data Notes

- All records are DCI-only and use `dci_work_unit`
- `status` supports `Open`, `In Review`, and `Closed`
- `closed_date` is set for closed requests

## Next Steps

- Add additional filters (status, assigned analyst)
- Add CSV export
- Integrate with SharePoint or Microsoft Lists when data model is finalized