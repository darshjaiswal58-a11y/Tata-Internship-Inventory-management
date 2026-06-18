# Tata Inventory Criteria Checker

Local MVP for daily Excel upload, stock criteria checking, and reorder report download.

## Run

```powershell
& "C:\Users\darsh\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" app.py
```

Then open:

```text
http://127.0.0.1:8000
```

## Demo Logins

- Admin: `admin@tatamotors.com` / `admin123`
- Employee: `employee@tatamotors.com` / `employee123`

Only the admin can view and edit criteria.

## Inventory Areas And Material Groups

The first dashboard has two inventory areas:

- Machinery Spare
- Tools

After selecting an area, use the Material Category dropdown to filter by part type. The dropdown is generated from:

```text
D:\tata_internship\updated\Material group 3002 and 3004.xlsx
```

The app uses the `Material Group` column as the material category source.

Daily Spare/Tools uploads can also include a category column. The app checks these names:

- `Material Group`
- `Material Category`
- `Part Type`
- `Category`

When one of those columns exists, uploaded rows update the dashboard material-category list automatically, including new categories not present in the original master file.

Default material critical values are loaded from:

```text
E:\downloads\Phase1_Critical_Stock_Analysis.xlsx
```

Tools critical values are loaded from:

```text
D:\tata_internship\updated\stockitemlist30043002plant\Machinery_Spares_Critical_Values_Final.xlsx
```

Additional Tools stock-list items are loaded from:

```text
D:\tata_internship\updated\stockitemlist30043002plant\final stock list 25-26 (MACHINERY SPARES) jan dark room.xlsx
```

Dashboard stock-zone counts are loaded from:

```text
E:\downloads\330_Eligible_Materials_Zone_Classification.xlsx
```

Dashboard stock-parts count is loaded from:

```text
E:\downloads\330_Unique_Eligible_Materials.xlsx
```

The app uses:

- `Critical Value` as the default material critical/minimum stock
- `Net_Consumption` or `Usage Frequency` as the material net consumption value shown in analysis and reports
- `Zone Status` as the Red / Yellow / Green dashboard count source

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
