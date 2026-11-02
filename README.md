# xlwings Wingman sales report demo

Build a reusable sales summary from a prompt, then rerun the saved Python script when the data changes.

## Files

- `wingman-practice.xlsx`: 12 fictional order lines in the named Excel Table `tblSales` on the Data worksheet, with filter buttons and a frozen header.
- `prompt.txt`: the exact prompt used for the walkthrough.

## Follow along

1. Install [xlwings Lite](https://lite.xlwings.org/quickstart) from Excel's Add-ins store and open the practice workbook. Save a working copy.
2. Open Wingman and connect your own supported AI provider. If you were personally given a beta code, use Wingman's **+ > Enter beta code** option. This download contains no API key or credits.
3. Leave access set to **Edit workbook**. Paste `prompt.txt` into Wingman and submit it.
4. Inspect the saved Python code in the editor. Verify the Sales Summary: East 34 units / $529; West 23 units / $342; grand total 57 units / $871.
5. Change `Data!D2` from 10 to 11. Select **Refresh Sales Report** in the script dropdown and click Run. The new total should be 58 units / $883, with East at $541. This action runs the saved script without sending another chat prompt.
6. To check new rows, enter `S013`, `South`, `Mug`, `2`, `20` in `Data!A14:E14` and rerun. With D2 still 11, expect 60 units / $923. Restore D2 to 10 and clear A14:E14 when finished, then rerun.

Keep headers unchanged and order rows contiguous. Confirm the Table name under Table Design. Add new orders inside the Table; if pasting below it does not expand it, use Table Design > Resize Table. The prompt and saved calculation in the walkthrough read the contiguous Data!A:E source range. The example uses whole-number units. A script-generated report is refreshed by running its script; changing an input does not by itself refresh written output values. AI output can vary, so check the code and totals you receive.

Tested in desktop Excel on Windows with Wingman beta access on November 2, 2026. Provider availability and costs can change. See the [Wingman documentation](https://lite.xlwings.org/wingman) and [security notes](https://lite.xlwings.org/security#wingman).
