# Process Atlas frontend

Interactive process and data ontology navigator for Pharmaceuticals, Food & Beverage, and Oil & Gas.

## Run locally

```bash
npm install
npm run dev
```

Use `npm run build`, `npm run lint`, and `npm test` for validation.

## Run the production build

```bash
npm run build
npm run preview
```

The preview server listens on all interfaces and defaults to `http://localhost:4173`.

## Deploy with Docker

The multi-stage `Dockerfile` builds the application and serves it through Nginx with SPA fallback, immutable asset caching, security headers, and an HTTP health check.

```bash
docker build -t process-atlas .
docker run --rm -p 8080:8080 process-atlas
```

Open `http://localhost:8080`. The container health check probes the same endpoint.

## Deploy to static hosting

Run `npm ci && npm run build` in CI and publish the generated `dist` directory. Configure the hosting platform to rewrite unknown routes to `index.html`. No server-side runtime or environment variables are required.

## Excel data sources

Each industry has its own workbook under `public/data`:

- `pharma.xlsx`
- `food-beverage.xlsx`
- `oil-gas.xlsx`

All workbooks use the same template. Required sheets are:

- `Overview`
- `L1 - Stages`
- `L2 - Processes`
- `L3 - Sub-steps`
- `L3 - Process Steps`
- `Functions & Capabilities`
- `Entity Catalog`
- `Entity-to-L3 Mapping`
- `Entity Relationships`

The application derives business-function and capability navigation from the workbook taxonomy and L3 ownership columns. Entity ownership comes from `Entity-to-L3 Mapping`; graph links come from `Entity Relationships`.

The predefined workbook in `public/data` is the runtime source of truth. Selecting an industry fetches and parses its workbook with cache disabled, so updates to the workbook are reflected after a page refresh or when the sector is selected again; no manual import is required. Use **Import Excel** to preview a local workbook in the current session and **Export Excel** to download the current model in the same template. Imported data may be retained in browser local storage for the current session, but it never overrides the predefined workbook on startup.

Run `npm run generate:data` to regenerate the Food & Beverage and Oil & Gas baseline workbooks from the fallback catalog. The supplied life-sciences workbook remains the authoritative bundled `pharma.xlsx` and is intentionally excluded from regeneration.
