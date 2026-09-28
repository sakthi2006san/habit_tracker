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

The workbook and configured sheet are created on first run. Each habit is stored as a row, with completion history and custom days encoded as JSON in their columns. A `Daily Log` sheet also lists each completed habit by date for day-by-day review. Keep the local API running while using the app; the browser's local storage remains a cache and fallback.

The API can also be started separately with `npm run server`.
