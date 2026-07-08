# Illinois State Police FOIA Command Dashboard

Static dashboard prototype for tracking Illinois State Police FOIA requests.

This version uses a local JSON file (`foia-data.json`) as a temporary data source.
No SharePoint or Microsoft Lists integration is included yet.

## Files

- `index.html` - dashboard layout and UI containers
- `style.css` - executive-style law enforcement dashboard theme
- `app.js` - data loading, KPI calculations, table rendering, and interactions
- `foia-data.json` - sample FOIA request records

## Features

- Header: Illinois State Police / FOIA Command Dashboard
- KPI cards:
	- Total Open FOIAs
	- Due in 10 Days
	- Due in 5 Days
	- Due Today
	- Overdue
	- Average Days Open
	- Closed This Month
	- Requests by Division
- Division workload table (DCI, DJS, DOP, DII, DFS)
- Upcoming deadlines panel
- Searchable FOIA request table
- Right-side detail panel for selected request
- Dark mode toggle
- Presentation mode toggle

## Run Locally

Because the dashboard fetches JSON, run it from a local web server.

Option 1 (Python):

```bash
python3 -m http.server 8080
```

Then open:

`http://localhost:8080`

## Data Notes

- Division values are constrained to: `DCI`, `DJS`, `DOP`, `DII`, `DFS`
- `status` supports `Open`, `In Review`, and `Closed`
- `closed_date` is set for closed requests

## Next Steps

- Add filters (division, status, assigned analyst)
- Add CSV export
- Integrate with SharePoint or Microsoft Lists when data model is finalized