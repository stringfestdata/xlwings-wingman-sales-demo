import xlwings as xw
from xlwings import script


SALES_SHEET_NAME = "Data"
SALES_TABLE_NAME = "tblSales"
SUMMARY_SHEET_NAME = "Sales Summary"
SUMMARY_TABLE_NAME = "tblSalesSummary"
INPUT_ORDER_ID = "Order ID"
INPUT_REGION = "Region"
INPUT_PRODUCT = "Product"
INPUT_UNITS = "Units"
INPUT_UNIT_PRICE = "Unit price (USD)"
INPUT_HEADERS = [
    INPUT_ORDER_ID,
    INPUT_REGION,
    INPUT_PRODUCT,
    INPUT_UNITS,
    INPUT_UNIT_PRICE,
]
SUMMARY_HEADERS = ["Region", "Units", "Revenue (USD)"]
HEADER_FILL = "#CF3338"
CHARCOAL = "#1F2120"
WHITE = "#FFFFFF"
TOTAL_LABEL = "Total"
UNITS_TOTAL_FORMULA = "=SUBTOTAL(109,tblSalesSummary[Units])"
REVENUE_TOTAL_FORMULA = "=SUBTOTAL(109,tblSalesSummary[Revenue (USD)])"
SUMMARY_UNITS_COLUMN = "Units"
SUMMARY_REVENUE_COLUMN = "Revenue (USD)"


def is_whole_units(value: object) -> bool:
    """Return True for a whole, nonnegative number (not a bool or text)."""
    if isinstance(value, bool):
        return False
    if isinstance(value, int):
        return value >= 0
    if isinstance(value, float):
        return value.is_integer() and value >= 0
    return False


def is_nonnegative_number(value: object) -> bool:
    """Return True for a numeric, nonnegative value (not a bool or text)."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    return value >= 0


def is_blank(value: object) -> bool:
    """Return True when a cell is empty or whitespace-only text."""
    return value is None or str(value).strip() == ""


def summarize_regions(
    header: list[str], rows: list[list[object]]
) -> list[tuple[str, int, float]]:
    """Validate tblSales rows and return (region, units, revenue) sorted by region.

    Column positions are looked up by header name, never by fixed position.
    """
    missing = [name for name in INPUT_HEADERS if name not in header]
    if missing:
        raise ValueError(f"tblSales is missing required column(s): {missing}")
    if not rows:
        raise ValueError("tblSales has no data rows; output was not changed.")
    index = {name: header.index(name) for name in INPUT_HEADERS}
    problems: list[str] = []
    totals: dict[str, list[object]] = {}
    for offset, row in enumerate(rows):
        row_number = offset + 1
        order_id = row[index[INPUT_ORDER_ID]]
        region = row[index[INPUT_REGION]]
        units = row[index[INPUT_UNITS]]
        price = row[index[INPUT_UNIT_PRICE]]
        label = f"data row {row_number} (Order ID {order_id!r})"
        if is_blank(order_id):
            problems.append(f"{label}: blank Order ID")
        if is_blank(region):
            problems.append(f"{label}: blank Region")
        if not is_whole_units(units):
            problems.append(
                f"{label}: Units must be a whole number >= 0, got {units!r}"
            )
        if not is_nonnegative_number(price):
            problems.append(
                f"{label}: Unit price (USD) must be a number >= 0, got {price!r}"
            )
        if problems and problems[-1].startswith(label):
            continue
        region_name = str(region)
        key = region_name.casefold()
        if key not in totals:
            totals[key] = [region_name, 0, 0.0]
        totals[key][1] += int(units)
        totals[key][2] += float(units) * float(price)
    if problems:
        shown = "; ".join(problems[:20])
        raise ValueError(f"Validation failed; output was not changed. {shown}")
    summary = [tuple(item) for item in totals.values()]
    summary.sort(key=lambda item: item[0].casefold())
    return summary


def get_or_add_sheet(book: xw.BookAsync, sheet_name: str) -> xw.Sheet:
    """Return the named sheet, creating it when it does not exist."""
    for sheet in book.sheets:
        if sheet.name == sheet_name:
            return sheet
    return book.sheets.add(name=sheet_name)


def find_table(sheet: xw.Sheet, table_name: str) -> xw.Table | None:
    """Return the named Table on a sheet, or None."""
    for table in sheet.tables:
        if table.name == table_name:
            return table
    return None


@script(name="Refresh Sales Report")
async def refresh_sales_report(book: xw.BookAsync):
    """# Refresh Sales Report
    Summarizes tblSales by Region into tblSalesSummary on the Sales Summary sheet.
    """
    sales_sheet = book.sheets[SALES_SHEET_NAME]
    sales_table = sales_sheet.tables[SALES_TABLE_NAME]
    header_rows = await sales_table.header_row_range.options(ndim=2).get_value()
    header = [
        str(value).strip() if value is not None else "" for value in header_rows[0]
    ]
    rows = []
    body_range = sales_table.data_body_range
    if body_range is not None:
        rows = await body_range.options(ndim=2).get_value()
    summary = summarize_regions(header, rows)

    summary_sheet = get_or_add_sheet(book, SUMMARY_SHEET_NAME)
    summary_table = find_table(summary_sheet, SUMMARY_TABLE_NAME)
    body_values = [[region, units, revenue] for region, units, revenue in summary]
    body_count = len(body_values)
    old_footprint = None
    if summary_table is None:
        # Origin is used only to place the output Table on first creation.
        origin = summary_sheet.range((1, 1))
        seed = origin.resize(len(body_values) + 1, len(SUMMARY_HEADERS))
        seed.options(ndim=2).value = [SUMMARY_HEADERS] + body_values
        summary_table = summary_sheet.tables.add(
            source=seed,
            name=SUMMARY_TABLE_NAME,
            table_style_name="TableStyleLight1",
        )
        header_rng = seed.resize(1, len(SUMMARY_HEADERS))
    else:
        # Preserve the existing Table; clearing an intersecting range removes it in Lite.
        summary_table.show_totals = False
        old_footprint = summary_table.range
        # The Table's top-left cell is its header row; resize from its own range.
        header_rng = old_footprint.resize(1, len(SUMMARY_HEADERS))
        summary_table.resize(old_footprint.resize(body_count + 1, len(SUMMARY_HEADERS)))
        header_rng.options(ndim=2).value = [SUMMARY_HEADERS]
        header_rng.offset(1, 0).resize(body_count, len(SUMMARY_HEADERS)).options(
            ndim=2
        ).value = body_values

    summary_table.show_totals = True
    # Totals row sits directly below the body inside the Table footprint.
    totals_rng = header_rng.offset(body_count + 1, 0).resize(1, len(SUMMARY_HEADERS))
    totals_rng.formula = [[TOTAL_LABEL, UNITS_TOTAL_FORMULA, REVENUE_TOTAL_FORMULA]]

    footprint = header_rng.resize(body_count + 2, len(SUMMARY_HEADERS))
    footprint.color = WHITE
    footprint.font.color = CHARCOAL
    header_rng.color = HEADER_FILL
    header_rng.font.set(bold=True, color=WHITE)
    totals_rng.font.bold = True
    totals_rng.offset(0, SUMMARY_HEADERS.index(SUMMARY_UNITS_COLUMN)).resize(1, 1).number_format = "#,##0"
    totals_rng.offset(0, SUMMARY_HEADERS.index(SUMMARY_REVENUE_COLUMN)).resize(1, 1).number_format = "$#,##0.00"
    body_rng = header_rng.offset(1, 0).resize(body_count, len(SUMMARY_HEADERS))
    body_rng.offset(0, 1).resize(body_count, 1).number_format = "#,##0"
    body_rng.offset(0, 2).resize(body_count, 1).number_format = "$#,##0.00"
    footprint.columns.autofit()
    # Clear only obsolete rows below the resized Table, after the queued resize.
    if old_footprint is not None and old_footprint.rows.count > body_count + 2:
        old_footprint.offset(body_count + 2, 0).resize(
            old_footprint.rows.count - body_count - 2, len(SUMMARY_HEADERS)
        ).clear()

    grand_units = sum(item[1] for item in summary)
    grand_revenue = sum(item[2] for item in summary)
    print(f"Input rows: {len(rows)}")
    print(f"Regions: {len(summary)}")
    print(f"Units total: {grand_units}")
    print(f"Revenue total (USD): {grand_revenue:,.2f}")
