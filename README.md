# APCRDA Local Finance Portfolio Website

This repository now contains a localhost-ready website that exhibits the four Python dashboard ideas as one browser tool:

1. `project1_uc_dashboard.py` → UC compliance and next-tranche risk.
2. `project2_contractor_aging.py` → contractor bill aging, MSME interest, and stall risk.
3. `project3_budget_velocity.py` → budget utilization velocity and lapse risk.
4. `project4_ppp_liability.py` → PPP obligations, guarantees, revenue share, and disclosure gaps.

## Run locally

```bash
npm start
```

Then open <http://localhost:4173>.

No npm packages are required. `website/server.js` uses Node's built-in HTTP server and serves the static files in `website/`.

## File system

```text
package.json                         # localhost start command
website/
  server.js                          # zero-dependency local web server
  index.html                         # single-page dashboard shell
  styles.css                         # responsive presentation styling
  app.js                             # filtering, charts, table rendering, CSV export
  data/dashboard-data.json           # exhibit dataset for all four tools
```

## Replacing the exhibit dataset

The current dataset is synthetic and presentation-ready. To use real registers, keep the same JSON shape in `website/data/dashboard-data.json`:

- `tools[]` contains one object per dashboard.
- `kpis[]` drives the KPI cards.
- `records[]` drives search, filtering, charts, and CSV download.
- `statusField` and `amountField` tell the app which record columns to chart.

