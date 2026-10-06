# TrimWorks Estimator

A free estimating platform for **interior trim, millwork, and doors** contractors.

## What it does
- Create unlimited estimates for clients
- **Doors** — pick door types (hollow/solid core, prehung, barn, pocket, bifold…), sizes, hardware, and labor hours
- **Trim & molding** — priced by linear foot with automatic waste factor
- **Custom millwork** — built-ins, shelving, wainscoting, stair parts
- Automatic totals: materials, labor, markup/overhead & profit, and sales tax
- **Price catalog** — set your own default supplier prices and labor rates
- **Print / Save as PDF** — clean, client-ready estimate with your company info and a signature line
- Track estimate status: Draft, Sent, Won, Lost
- Backup & restore all your data with one click

## How to use it
Just open the website — no installation, no account, no coding. Your data is saved
automatically in your browser. Use **Backup Data** regularly to keep a safe copy.

1. Go to **My Company** and enter your business info (shows on printed estimates)
2. Go to **Price Catalog** and adjust prices to match your suppliers
3. Click **+ New Estimate** and start adding doors, trim, and millwork
4. Click **Print / Save as PDF** to give the estimate to your client

## Forms
- **Window & Door Monthly RPA Review** — `forms/rpa/Four_Corners_Building_Supply_RPA_Window_Door.pdf`.
  Fillable one-page Results / Pipeline / Activity review for the Window & Door department. Same layout
  as the company Monthly RPA Review (Pipeline is yearly), plus an **Orders In & Quotes** section: orders in / written business
  and quoted dollars for the month vs. budget. To change it, edit `forms/rpa/build_rpa_window_door.py`
  and run `python3 forms/rpa/build_rpa_window_door.py` (needs `pip install reportlab`).
