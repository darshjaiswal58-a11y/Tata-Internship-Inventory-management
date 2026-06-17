# Tata Inventory Criteria Checker

Local MVP for daily Excel upload, stock criteria checking, and reorder report download.

## Run

```powershell
& "C:\Users\darsh\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" app.py
```

Then open:

```text
http://localhost:8000
```

## Demo Logins

- Admin: `admin@tatamotors.com` / `admin123`
- Employee: `employee@tatamotors.com` / `employee123`

Only the admin can view and edit criteria.

## Inventory Areas And Material Groups

The first dashboard has two inventory areas:

- M/L Spare
- Tools

After selecting an area, use the Material Category dropdown to filter by part type. The dropdown is generated from:

```text
D:\tata_internship\updated\Material group 3002 and 3004.xlsx
```

The app uses the `Material Group` column as the material category source.

## Required Upload Columns

- Purchase Order Date
- Entry Date
- Material
- Material Description
- Quantity
- Days Between
- Valuated Stock

Rules:

- Positive `Quantity` = received stock
- Negative `Quantity` = used stock
- Report includes items where `Valuated Stock <= Minimum Stock`
