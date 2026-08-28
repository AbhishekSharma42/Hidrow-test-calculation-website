# Abhi Sharma Hydro-Test Calculator

React + Vite website based on the supplied `calculations hydro.xlsx`.

Pages:
1. Dashboard
2. HT Section Introduction
3. Air Volume Calculation
4. Strength & Leak Test

All editable form fields start empty. Calculations update live as the user enters values.

## Run
```bash
npm install
npm run dev
```

Then open the Vite address shown in the terminal (normally http://localhost:5173).

## Important
The workbook contains large lookup tables and some Excel-specific array formulas. This web version reproduces the main user-facing calculations and interpolation behavior in JavaScript rather than embedding Excel itself.
