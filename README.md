# xlwings Wingman sales report demo

Build a reusable report from the named Excel Table `tblSales`, then test the saved script as the orders change.

## Files

- `wingman-practice.xlsx`: 12 synthetic order lines in `tblSales` on the Data sheet.
- `prompt.txt`: the walkthrough prompt, including the Table-reference requirements refined during testing.
- `refresh_sales_report.py`: the tested report code for comparison and correction in the walkthrough.

## Follow along

1. Install [xlwings Lite](https://lite.xlwings.org/quickstart) from Excel's Add-ins store. Open the workbook and save a working copy.
2. Open Wingman and connect a supported AI provider, or use a beta code if you were personally given one. No credentials or AI credits are included here.
3. Set access to **Edit workbook**, submit `prompt.txt`, and inspect the saved `refresh_sales_report` code. It should read `tblSales` by Table and column name, and create a native output Table named `tblSalesSummary` with a separate totals row. Compare it with `refresh_sales_report.py`. If needed, replace the contents of `main.py` in the xlwings Lite editor with that file and save. The tested code preserves the existing output Table and clears obsolete rows only after resizing.
4. Select **Refresh Sales Report** from the script dropdown and click **Run**. Check East: 34 units / $529; West: 23 units / $342; total: 57 units / $871. The source should be unchanged.
5. In `tblSales`, change **Units** for **Order ID S001** from 10 to 11. Run the saved script without sending another chat message. Expect 58 units / $883, with East at 35 units / $541.
6. Add a row inside `tblSales`: **Order ID** S013; **Region** South; **Product** Mug; **Units** 2; **Unit price (USD)** 20. With S001 still at 11 units, rerun and expect 60 units / $923, including South at 2 units / $40.
7. Restore **S001's Units** to 10. Select the S013 record and use **Home > Delete > Delete Table Rows**. Rerun and check that totals return to 57 units / $871, with no stale South row or duplicate output Table. Save the workbook.

Keep one order per row, unchanged column names, numeric quantities and prices, and no merged cells or subtotals inside the source Table. Confirm its name under **Table Design > Table Name**. To append a record, select the final order's **Unit price (USD)** field and press Tab.

Independent checks: `=SUM(tblSales[Units])` and `=SUMPRODUCT(tblSales[Units],tblSales[Unit price (USD)])`. Enter these outside the input Table. The report's data values refresh when you run the saved script; AI output can vary, so inspect the code and reconcile the totals.

Tested in desktop Excel on Windows with Wingman beta access on November 2, 2026. See the [Wingman documentation](https://lite.xlwings.org/wingman) and [security notes](https://lite.xlwings.org/security#wingman).
