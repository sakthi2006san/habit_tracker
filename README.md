# Momentum Habit Tracker

## Run locally

```sh
npm install
npm run dev
```

The development command starts the Vite app and its local Excel storage API. Configure the workbook and worksheet in `application.properties`:

```properties
excel.file.path=./data/habits.xlsx
excel.sheet.name=Habits
```

The workbook and configured sheet are created on first run. Each habit is stored as a row, with completion history and custom days encoded as JSON in their columns. Completed habits are also listed by date in monthly sheets named `YYYY-MM`, such as `2026-09`. Keep the local API running while using the app; the browser's local storage remains a cache and fallback.

The API exposes `GET /api/health`, `GET /api/config`, and `GET`/`PUT /api/habits`. A successful habit save is confirmed by reading the updated history back from the workbook; malformed requests return `400`, while workbook failures return `500`.

The API can also be started separately with `npm run server`.
