"""
Step 10 - Generate the Power BI project (PBIP) for the Supply Chain Control Tower.

Creates powerbi/SupplyChainControlTower.pbip with:
  * SupplyChainControlTower.SemanticModel  (TMDL: tables, Power Query, relationships,
                                            calculated columns, 50+ DAX measures)
  * SupplyChainControlTower.Report         (PBIR: 5 report pages with visuals)

Open the .pbip file in Power BI Desktop, click Refresh, then File > Save as .pbix if
you want a single file.

Usage:
    python src/build_powerbi.py "C:\\path\\to\\supply-chain-inventory-risk-analytics\\data\\"
The argument is where Power BI should read the CSV files (the project's data folder).
"""
import json
import shutil
import sys
import uuid
from pathlib import Path

import config

NAME = "SupplyChainControlTower"
PB = config.ROOT / "powerbi"
SM = PB / f"{NAME}.SemanticModel"
RP = PB / f"{NAME}.Report"
TEMPLATE = PB / "_template"


def tag() -> str:
    """Stable lineage tags, so re-running the script doesn't create noisy git diffs."""
    tag.counter += 1
    return str(uuid.uuid5(uuid.NAMESPACE_URL, f"{NAME}/{tag.counter}"))


tag.counter = 0


def q(name: str) -> str:
    """Quote a TMDL object name if needed."""
    return name if name.replace("_", "").isalnum() and not name[0].isdigit() else "'" + name.replace("'", "''") + "'"


def indent_block(text: str, tabs: int) -> str:
    pad = "\t" * tabs
    return "\n".join(pad + line if line.strip() else "" for line in text.strip("\n").splitlines())


# ---------------------------------------------------------------------------
# Model definition
# ---------------------------------------------------------------------------
def m_csv(file: str, types: list[tuple[str, str]]) -> str:
    type_list = ", ".join(f'{{"{c}", {t}}}' for c, t in types)
    return f'''let
    Source = Csv.Document(File.Contents(DataFolder & "{file}"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Headers, {{{type_list}}}, "en-US")
in
    Typed'''


M_TYPE = {"string": "type text", "double": "type number", "int64": "Int64.Type", "dateTime": "type date"}

# import tables: name -> (csv, [(column, dataType, formatString or None)])
IMPORT_TABLES = {
    "products": ("products.csv", [("sku", "string", None), ("product_name", "string", None), ("category", "string", None),
                                  ("unit_cost_cad", "double", "$#,0.00"), ("supplier_id", "string", None)]),
    "suppliers": ("suppliers.csv", [("supplier_id", "string", None), ("supplier_name", "string", None),
                                    ("country", "string", None), ("promised_lead_time_wks", "int64", "0")]),
    "weekly_demand": ("weekly_demand.csv", [("week_start", "dateTime", "yyyy-mm-dd"), ("sku", "string", None),
                                            ("units", "int64", "#,0")]),
    "purchase_orders": ("purchase_orders.csv", [
        ("po_id", "string", None), ("supplier_id", "string", None), ("sku", "string", None),
        ("order_date", "dateTime", "yyyy-mm-dd"), ("promised_date", "dateTime", "yyyy-mm-dd"),
        ("received_date", "dateTime", "yyyy-mm-dd"), ("qty_ordered", "int64", "#,0"), ("qty_received", "int64", "#,0"),
        ("qty_defective", "int64", "#,0"), ("po_value_cad", "double", "$#,0")]),
    "inventory_snapshot": ("inventory_snapshot.csv", [("sku", "string", None), ("on_hand_units", "int64", "#,0"),
                                                      ("on_order_units", "int64", "#,0")]),
    "landed_cost_by_sku": ("landed_cost_by_sku.csv", [
        ("sku", "string", None), ("product_name", "string", None), ("abc_class", "string", None),
        ("current_supplier", "string", None), ("current_country", "string", None),
        ("nearshore_supplier", "string", None), ("nearshore_country", "string", None),
        ("annual_units", "double", "#,0"), ("current_landed_unit_cad", "double", "$#,0.00"),
        ("nearshore_landed_unit_cad", "double", "$#,0.00"), ("current_lead_time_wks", "double", "0.0"),
        ("nearshore_lead_time_wks", "double", "0.0"), ("current_total_cad", "double", "$#,0"),
        ("nearshore_total_cad", "double", "$#,0"), ("annual_saving_cad", "double", "$#,0"),
        ("saving_pct", "double", "0.0%"), ("inventory_released_cad", "double", "$#,0"),
        ("recommendation", "string", None)]),
    "landed_cost_by_country": ("landed_cost_by_country.csv", [
        ("country", "string", None), ("region", "string", None), ("price", "double", "0.0"),
        ("freight", "double", "0.0"), ("duty", "double", "0.0"), ("brokerage", "double", "0.0"),
        ("landed_total", "double", "0.0"), ("avg_lead_time_wks", "double", "0.0"),
        ("on_time_rate", "double", "0.0%")]),
}

# calculated columns: table -> [(name, dataType, formatString, expression)]
CALC_COLUMNS = {
    "purchase_orders": [
        ("Lead Time Wks", "double", "0.0", "DATEDIFF ( purchase_orders[order_date], purchase_orders[received_date], DAY ) / 7"),
        ("On Time", "int64", "0", "IF ( purchase_orders[received_date] <= purchase_orders[promised_date], 1, 0 )"),
        ("Days Late", "int64", "0", "MAX ( 0, INT ( purchase_orders[received_date] - purchase_orders[promised_date] ) )"),
    ],
    "suppliers": [
        ("Country Risk", "int64", "0", '''SWITCH ( suppliers[country],
    "Canada", 1, "USA", 2, "Germany", 2, "Mexico", 3,
    "Malaysia", 3, "India", 4, "Vietnam", 4, "China", 5,
    3 )'''),
    ],
    "products": [
        ("Avg Weekly Demand", "double", "#,0.0", '''VAR k = products[sku]
VAR lastWk = MAX ( weekly_demand[week_start] )
RETURN
    AVERAGEX (
        FILTER ( weekly_demand, weekly_demand[sku] = k && weekly_demand[week_start] > lastWk - 364 ),
        weekly_demand[units] )'''),
        ("Demand Std", "double", "#,0.0", '''VAR k = products[sku]
VAR lastWk = MAX ( weekly_demand[week_start] )
RETURN
    STDEVX.S (
        FILTER ( weekly_demand, weekly_demand[sku] = k && weekly_demand[week_start] > lastWk - 364 ),
        weekly_demand[units] )'''),
        ("Annual Units", "double", "#,0", "products[Avg Weekly Demand] * 52"),
        ("Annual Spend", "double", "$#,0", "products[Annual Units] * products[unit_cost_cad]"),
        ("LT Mean Wks", "double", "0.0", '''VAR s = products[supplier_id]
RETURN AVERAGEX ( FILTER ( purchase_orders, purchase_orders[supplier_id] = s ), purchase_orders[Lead Time Wks] )'''),
        ("LT Std Wks", "double", "0.00", '''VAR s = products[supplier_id]
RETURN STDEVX.S ( FILTER ( purchase_orders, purchase_orders[supplier_id] = s ), purchase_orders[Lead Time Wks] )'''),
        ("On Hand", "int64", "#,0", "LOOKUPVALUE ( inventory_snapshot[on_hand_units], inventory_snapshot[sku], products[sku] )"),
        ("On Order", "int64", "#,0", "LOOKUPVALUE ( inventory_snapshot[on_order_units], inventory_snapshot[sku], products[sku] )"),
        ("Inventory Position", "int64", "#,0", "products[On Hand] + products[On Order]"),
        ("Cum Share Before", "double", "0.0%", '''VAR me = products[Annual Spend]
RETURN DIVIDE ( SUMX ( FILTER ( products, products[Annual Spend] > me ), products[Annual Spend] ), SUM ( products[Annual Spend] ) )'''),
        ("ABC", "string", None, 'IF ( products[Cum Share Before] < 0.80, "A", IF ( products[Cum Share Before] < 0.95, "B", "C" ) )'),
        ("Demand CV", "double", "0.00", "DIVIDE ( products[Demand Std], products[Avg Weekly Demand] )"),
        ("XYZ", "string", None, 'IF ( products[Demand CV] <= 0.25, "X", IF ( products[Demand CV] <= 0.50, "Y", "Z" ) )'),
        ("Segment", "string", None, "products[ABC] & products[XYZ]"),
        ("EOQ", "int64", "#,0", '''VAR S = 85
VAR H = products[unit_cost_cad] * 0.22
RETURN IF ( H > 0, ROUNDUP ( SQRT ( 2 * products[Annual Units] * S / H ), 0 ), 0 )'''),
    ],
}

MEASURES = [
    # ---- inventory policy
    ("z Score", "0.000", "NORM.S.INV ( [Service Level Value] )"),
    ("Safety Stock", "#,0", '''VAR z = [z Score]
RETURN
    SUMX ( products,
        ROUNDUP ( z * SQRT ( products[LT Mean Wks] * products[Demand Std] ^ 2
                           + products[Avg Weekly Demand] ^ 2 * products[LT Std Wks] ^ 2 ), 0 ) )'''),
    ("Reorder Point", "#,0", "SUMX ( products, ROUNDUP ( products[Avg Weekly Demand] * products[LT Mean Wks] + [Safety Stock], 0 ) )"),
    ("Inventory Status", None, '''IF ( HASONEVALUE ( products[sku] ),
    VAR pos = SELECTEDVALUE ( products[Inventory Position] )
    VAR ss  = [Safety Stock]
    VAR rop = [Reorder Point]
    VAR mx  = rop + SELECTEDVALUE ( products[EOQ] )
    RETURN
        SWITCH ( TRUE (),
            pos < ss,   "Critical - expedite",
            pos <= rop, "Order now",
            pos > mx,   "Overstock",
            "Healthy" ) )'''),
    ("Suggested Order Qty", "#,0", '''SUMX ( products,
    VAR st = [Inventory Status]
    VAR gap = [Reorder Point] + products[EOQ] - products[Inventory Position]
    RETURN IF ( st IN { "Critical - expedite", "Order now" }, MAX ( products[EOQ], gap ), 0 ) )'''),
    ("SKUs in Status", "0", "COUNTROWS ( FILTER ( VALUES ( products[sku] ), [Inventory Status] = SELECTEDVALUE ( 'Status List'[Status] ) ) ) + 0"),
    ("SKUs to Order", "0", 'COUNTROWS ( FILTER ( VALUES ( products[sku] ), [Inventory Status] IN { "Critical - expedite", "Order now" } ) ) + 0'),
    ("SKUs Critical", "0", 'COUNTROWS ( FILTER ( VALUES ( products[sku] ), [Inventory Status] = "Critical - expedite" ) ) + 0'),
    ("SKUs Overstocked", "0", 'COUNTROWS ( FILTER ( VALUES ( products[sku] ), [Inventory Status] = "Overstock" ) ) + 0'),
    ("Excess Inventory Value", "$#,0", '''SUMX ( products,
    VAR mx = [Reorder Point] + products[EOQ]
    RETURN IF ( products[Inventory Position] > mx, ( products[Inventory Position] - mx ) * products[unit_cost_cad], 0 ) )'''),
    ("Safety Stock Value", "$#,0", "SUMX ( products, [Safety Stock] * products[unit_cost_cad] )"),
    ("Annual Spend Total", "$#,0", "SUM ( products[Annual Spend] )"),
    ("A-Class Share of Spend", "0.0%", 'DIVIDE ( CALCULATE ( [Annual Spend Total], products[ABC] = "A" ), CALCULATE ( [Annual Spend Total], ALL ( products[ABC] ) ) )'),
    ("EOQ Annual Saving", "$#,0", '''SUMX ( products,
    VAR D = products[Annual Units]
    VAR H = products[unit_cost_cad] * 0.22
    VAR Qnow = products[Avg Weekly Demand] * 4
    VAR Qeoq = products[EOQ]
    VAR costNow = DIVIDE ( D, Qnow ) * 85 + Qnow / 2 * H
    VAR costEoq = DIVIDE ( D, Qeoq ) * 85 + Qeoq / 2 * H
    RETURN costNow - costEoq )'''),
    ("SKU Count", "0", "COUNTROWS ( products )"),
    # ---- supplier scorecard
    ("PO Count", "#,0", "COUNTROWS ( purchase_orders )"),
    ("PO Spend", "$#,0", "SUM ( purchase_orders[po_value_cad] )"),
    ("Share of Spend", "0.0%", "DIVIDE ( [PO Spend], CALCULATE ( [PO Spend], ALL ( suppliers ) ) )"),
    ("On-Time Delivery %", "0.0%", "AVERAGE ( purchase_orders[On Time] )"),
    ("Avg Days Late", "0.0", "AVERAGE ( purchase_orders[Days Late] )"),
    ("Avg Lead Time Wks", "0.0", "AVERAGE ( purchase_orders[Lead Time Wks] )"),
    ("Lead Time CV", "0.0%", "DIVIDE ( STDEV.S ( purchase_orders[Lead Time Wks] ), [Avg Lead Time Wks] )"),
    ("Fill Rate", "0.0%", "DIVIDE ( SUM ( purchase_orders[qty_received] ), SUM ( purchase_orders[qty_ordered] ) )"),
    ("Defect Rate", "0.00%", "DIVIDE ( SUM ( purchase_orders[qty_defective] ), SUM ( purchase_orders[qty_received] ) )"),
    ("Risk Score", "0", '''VAR fLate    = MIN ( 1, MAX ( 0, ( 1 - [On-Time Delivery %] ) / 0.35 ) )
VAR fLTV     = MIN ( 1, MAX ( 0, [Lead Time CV] / 0.30 ) )
VAR fQuality = MIN ( 1, MAX ( 0, [Defect Rate] / 0.025 ) )
VAR fCountry = ( AVERAGE ( suppliers[Country Risk] ) - 1 ) / 4
VAR fSpend   = MIN ( 1, MAX ( 0, [Share of Spend] / 0.30 ) )
RETURN
    IF ( HASONEVALUE ( suppliers[supplier_id] ) && [PO Count] > 0,
        100 * ( 0.30 * fLate + 0.20 * fLTV + 0.20 * fQuality + 0.15 * fCountry + 0.15 * fSpend ) )'''),
    ("Risk Tier", None, '''VAR s = [Risk Score]
RETURN IF ( ISBLANK ( s ), BLANK (), IF ( s > 60, "High", IF ( s > 35, "Medium", "Low" ) ) )'''),
    ("Main Risk Driver", None, '''VAR c1 = 0.30 * MIN ( 1, MAX ( 0, ( 1 - [On-Time Delivery %] ) / 0.35 ) )
VAR c2 = 0.20 * MIN ( 1, MAX ( 0, [Lead Time CV] / 0.30 ) )
VAR c3 = 0.20 * MIN ( 1, MAX ( 0, [Defect Rate] / 0.025 ) )
VAR c4 = 0.15 * ( AVERAGE ( suppliers[Country Risk] ) - 1 ) / 4
VAR c5 = 0.15 * MIN ( 1, MAX ( 0, [Share of Spend] / 0.30 ) )
VAR biggest = MAX ( MAX ( MAX ( c1, c2 ), MAX ( c3, c4 ) ), c5 )
RETURN
    IF ( HASONEVALUE ( suppliers[supplier_id] ) && [PO Count] > 0,
        SWITCH ( TRUE (),
            biggest = c1, "Late delivery",
            biggest = c2, "Lead-time variability",
            biggest = c3, "Quality",
            biggest = c4, "Country risk",
            "Spend dependency" ) )'''),
    ("High-Risk Suppliers", "0", "COUNTROWS ( FILTER ( VALUES ( suppliers[supplier_id] ), [Risk Score] > 60 ) ) + 0"),
    ("Risk Color", None, 'SWITCH ( [Risk Tier], "High", "#D03B3B", "Medium", "#FAB219", "Low", "#0CA30C" )'),
    ("Status Color", None, '''SWITCH ( [Inventory Status],
    "Critical - expedite", "#D03B3B", "Order now", "#EC835A", "Overstock", "#FAB219", "Healthy", "#0CA30C" )'''),
    ("Status Bar Color", None, '''SWITCH ( SELECTEDVALUE ( 'Status List'[Status] ),
    "Critical - expedite", "#D03B3B", "Order now", "#EC835A", "Overstock", "#FAB219", "Healthy", "#0CA30C" )'''),
    # ---- demand
    ("Units Sold", "#,0", "SUM ( weekly_demand[units] )"),
    # ---- disruption simulator
    ("Weeks of Stock", "0.0", "DIVIDE ( SUM ( products[Inventory Position] ), SUM ( products[Avg Weekly Demand] ) )"),
    ("New Lead Time Wks", "0.0", "IF ( NOT ISEMPTY ( products ), AVERAGE ( products[LT Mean Wks] ) + [Delay Weeks Value] )"),
    ("Units Short", "#,0", '''SUMX ( products,
    VAR wks = DIVIDE ( products[Inventory Position], products[Avg Weekly Demand] )
    VAR gap = MAX ( 0, products[LT Mean Wks] + [Delay Weeks Value] - wks )
    RETURN ROUNDUP ( gap * products[Avg Weekly Demand], 0 ) )'''),
    ("Revenue at Risk", "$#,0", "SUMX ( products, [Units Short] * products[unit_cost_cad] * 1.35 )"),
    ("SKUs Stocking Out", "0", "COUNTROWS ( FILTER ( VALUES ( products[sku] ), [Units Short] > 0 ) ) + 0"),
    ("Extra Buffer Cost", "$#,0", "SUMX ( products, ROUNDUP ( products[Avg Weekly Demand] * [Delay Weeks Value], 0 ) * products[unit_cost_cad] )"),
    ("Disruption Outcome", None, 'IF ( HASONEVALUE ( products[sku] ), IF ( [Units Short] > 0, "Stockout", "Covered" ) )'),
    ("Outcome Color", None, 'IF ( [Units Short] > 0, "#D03B3B", "#0CA30C" )'),
    # ---- logistics & landed cost (offshore vs nearshore)
    ("Offshore SKUs Reviewed", "0", "COUNTROWS ( landed_cost_by_sku ) + 0"),
    ("Nearshore Cheaper SKUs", "0",
     "COUNTROWS ( FILTER ( landed_cost_by_sku, landed_cost_by_sku[annual_saving_cad] > 0 ) ) + 0"),
    ("Nearshore Candidates", "0",
     'COUNTROWS ( FILTER ( landed_cost_by_sku, landed_cost_by_sku[recommendation] <> "Keep offshore" ) ) + 0'),
    ("Nearshore Saving", "$#,0",
     "SUMX ( FILTER ( landed_cost_by_sku, landed_cost_by_sku[annual_saving_cad] > 0 ), "
     "landed_cost_by_sku[annual_saving_cad] ) + 0"),
    ("Landed Saving", "$#,0", "SUM ( landed_cost_by_sku[annual_saving_cad] )"),
    ("Inventory Released", "$#,0",
     'CALCULATE ( SUM ( landed_cost_by_sku[inventory_released_cad] ), '
     'landed_cost_by_sku[recommendation] <> "Keep offshore" ) + 0'),
    ("Lead Time Cut Wks", "0.0",
     'AVERAGEX ( FILTER ( landed_cost_by_sku, landed_cost_by_sku[recommendation] <> "Keep offshore" ), '
     'landed_cost_by_sku[current_lead_time_wks] - landed_cost_by_sku[nearshore_lead_time_wks] )'),
    ("Price", "0.0", "SUM ( landed_cost_by_country[price] )"),
    ("Freight", "0.0", "SUM ( landed_cost_by_country[freight] )"),
    ("Duty", "0.0", "SUM ( landed_cost_by_country[duty] )"),
    ("Brokerage", "0.0", "SUM ( landed_cost_by_country[brokerage] )"),
    ("Landed Total", "0.0", "SUM ( landed_cost_by_country[landed_total] )"),
    ("Recommendation Color", None,
     'SWITCH ( SELECTEDVALUE ( landed_cost_by_sku[recommendation] ), '
     '"Nearshore: lower total cost", "#0CA30C", '
     '"Nearshore candidate: small premium for resilience", "#C98500", "#7A7974" )'),
]


def column_tmdl(name, dtype, fmt, source=None, expr=None, hidden=False, extra=()):
    head = f"\tcolumn {q(name)}"
    if expr is not None:
        head += " =\n" + indent_block(expr, 3)
    lines = [head, f"\t\tdataType: {dtype}"]
    if fmt:
        lines.append(f"\t\tformatString: {fmt}")
    if hidden:
        lines.append("\t\tisHidden")
    lines.append(f"\t\tlineageTag: {tag()}")
    numeric = dtype in ("int64", "double")
    lines.append(f"\t\tsummarizeBy: {'none' if not numeric or expr is not None or (source or '').startswith('[') else 'sum'}")
    if source:
        lines.append(f"\t\tsourceColumn: {source}")
    lines += [f"\t\t{e}" for e in extra]
    lines.append("")
    lines.append("\t\tannotation SummarizationSetBy = Automatic")
    return "\n".join(lines) + "\n"


def measure_tmdl(name, fmt, expr):
    head = f"\tmeasure {q(name)} =\n" + indent_block(expr, 3)
    lines = [head]
    if fmt:
        lines.append(f"\t\tformatString: {fmt}")
    lines.append(f"\t\tlineageTag: {tag()}")
    return "\n".join(lines) + "\n"


def partition_m(table, expr):
    return f"\tpartition {q(table)} = m\n\t\tmode: import\n\t\tsource =\n{indent_block(expr, 4)}\n"


def partition_calc(table, expr):
    return f"\tpartition {q(table)} = calculated\n\t\tmode: import\n\t\tsource =\n{indent_block(expr, 4)}\n"


def write(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def build_semantic_model(data_folder: str):
    d = SM / "definition"
    write(SM / "definition.pbism", json.dumps({"version": "4.2", "settings": {}}, indent=2))
    write(SM / ".platform", json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": str(uuid.uuid5(uuid.NAMESPACE_URL, NAME + "/sm"))}}, indent=2))
    write(d / "database.tmdl", "database\n\tcompatibilityLevel: 1606\n\n")
    write(d / "cultures" / "en-US.tmdl", "cultureInfo en-US\n\n")

    tables = list(IMPORT_TABLES) + ["Dates", "Status List", "Service Level", "Delay Weeks", "_Measures"]
    model = ["model Model", "\tculture: en-US", "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
             "\tsourceQueryCulture: en-US", "\tvalueFilterBehavior: independent", "\tdataAccessOptions",
             "\t\tlegacyRedirects", "\t\treturnErrorValuesAsNull", "",
             "annotation __PBI_TimeIntelligenceEnabled = 0", "",
             'annotation PBI_QueryOrder = ["DataFolder","' + '","'.join(IMPORT_TABLES) + '"]', "",
             'annotation PBI_ProTooling = ["DevMode"]', ""]
    model += [f"ref table {q(t)}" for t in tables]
    model += ["", "ref cultureInfo en-US", ""]
    write(d / "model.tmdl", "\n".join(model))

    folder = data_folder.replace('"', '""')
    write(d / "expressions.tmdl",
          f'expression DataFolder = "{folder}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n'
          f"\tlineageTag: {tag()}\n\n\tannotation PBI_ResultType = Text\n\n\tannotation PBI_NavigationStepName = Navigation\n\n")

    # import tables
    for t, (csv, cols) in IMPORT_TABLES.items():
        out = [f"table {q(t)}", f"\tlineageTag: {tag()}", ""]
        for name, dtype, fmt in cols:
            out.append(column_tmdl(name, dtype, fmt, source=name))
        for name, dtype, fmt, expr in CALC_COLUMNS.get(t, []):
            out.append(column_tmdl(name, dtype, fmt, expr=expr))
        out.append(partition_m(t, m_csv(csv, [(c, M_TYPE[dt]) for c, dt, _ in cols])))
        out.append("\tannotation PBI_ResultType = Table\n")
        write(d / "tables" / f"{t}.tmdl", "\n".join(out))

    # Dates (marked as date table)
    write(d / "tables" / "Dates.tmdl", "\n".join([
        "table Dates", "\tdataCategory: Time", f"\tlineageTag: {tag()}", "",
        column_tmdl("Date", "dateTime", "yyyy-mm-dd", source="[Date]", extra=("isKey", "isNameInferred")),
        column_tmdl("Year", "int64", "0", source="[Year]", extra=("isNameInferred",)),
        column_tmdl("Month", "string", None, source="[Month]", extra=("isNameInferred", "sortByColumn: 'Month Sort'")),
        column_tmdl("Month Sort", "int64", "0", source="[Month Sort]", hidden=True, extra=("isNameInferred",)),
        partition_calc("Dates", '''ADDCOLUMNS (
    CALENDAR ( MIN ( weekly_demand[week_start] ), MAX ( weekly_demand[week_start] ) + 90 ),
    "Year", YEAR ( [Date] ),
    "Month", FORMAT ( [Date], "mmm yyyy" ),
    "Month Sort", YEAR ( [Date] ) * 100 + MONTH ( [Date] ) )'''),
    ]))

    write(d / "tables" / "Status List.tmdl", "\n".join([
        "table 'Status List'", f"\tlineageTag: {tag()}", "",
        column_tmdl("Status", "string", None, source="[Status]", extra=("isNameInferred", "sortByColumn: Sort")),
        column_tmdl("Sort", "int64", "0", source="[Sort]", hidden=True, extra=("isNameInferred",)),
        partition_calc("Status List", '''DATATABLE (
    "Status", STRING, "Sort", INTEGER,
    { { "Critical - expedite", 1 }, { "Order now", 2 }, { "Healthy", 3 }, { "Overstock", 4 } } )'''),
    ]))

    write(d / "tables" / "Service Level.tmdl", "\n".join([
        "table 'Service Level'", f"\tlineageTag: {tag()}", "",
        measure_tmdl("Service Level Value", "0%", "SELECTEDVALUE ( 'Service Level'[Service Level %], 95 ) / 100"),
        column_tmdl("Service Level %", "int64", "0", source="[Value]"),
        partition_calc("Service Level", "GENERATESERIES ( 90, 99, 1 )"),
    ]))
    write(d / "tables" / "Delay Weeks.tmdl", "\n".join([
        "table 'Delay Weeks'", f"\tlineageTag: {tag()}", "",
        measure_tmdl("Delay Weeks Value", "0", "SELECTEDVALUE ( 'Delay Weeks'[Delay Weeks], 4 )"),
        column_tmdl("Delay Weeks", "int64", "0", source="[Value]"),
        partition_calc("Delay Weeks", "GENERATESERIES ( 0, 12, 1 )"),
    ]))

    out = ["table _Measures", f"\tlineageTag: {tag()}", ""]
    for name, fmt, expr in MEASURES:
        out.append(measure_tmdl(name, fmt, expr))
    out.append(column_tmdl("Value", "string", None, source="[Value]", hidden=True, extra=("isNameInferred",)))
    out.append(partition_calc("_Measures", '{ "" }'))
    write(d / "tables" / "_Measures.tmdl", "\n".join(out))

    rels = [("weekly_demand", "sku", "products", "sku"),
            ("products", "supplier_id", "suppliers", "supplier_id"),
            ("purchase_orders", "supplier_id", "suppliers", "supplier_id"),
            ("weekly_demand", "week_start", "Dates", "Date"),
            ("landed_cost_by_sku", "sku", "products", "sku")]
    rel_txt = []
    for ft, fc, tt, tc in rels:
        rel_txt.append(f"relationship {tag()}\n\tfromColumn: {q(ft)}.{q(fc)}\n\ttoColumn: {q(tt)}.{q(tc)}\n")
    write(d / "relationships.tmdl", "\n".join(rel_txt))
    write(SM / "diagramLayout.json", json.dumps({"version": "1.1.0", "diagrams": [{
        "ordinal": 0, "scrollPosition": {"x": 0, "y": 0}, "nodes": [], "name": "All tables", "zoomValue": 100,
        "pinKeyFieldsToTop": False, "showExtraHeaderInfo": False, "hideKeyFieldsWhenCollapsed": False,
        "tablesLocked": False}], "selectedDiagram": "All tables", "defaultDiagram": "All tables"}, indent=2))


# ---------------------------------------------------------------------------
# Report (PBIR)
# ---------------------------------------------------------------------------
SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"


def build_report(pages):
    write(RP / "definition.pbir", json.dumps({"version": "4.0", "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}}, indent=2))
    write(RP / ".platform", json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": str(uuid.uuid5(uuid.NAMESPACE_URL, NAME + "/rp"))}}, indent=2))
    base = "Fluent2-CY26SU09"
    res = RP / "StaticResources"
    (res / "SharedResources" / "BaseThemes").mkdir(parents=True, exist_ok=True)
    shutil.copy(TEMPLATE / f"{base}.json", res / "SharedResources" / "BaseThemes" / f"{base}.json")
    (res / "RegisteredResources").mkdir(parents=True, exist_ok=True)
    shutil.copy(PB / "SupplyChainTheme.json", res / "RegisteredResources" / "SupplyChainTheme.json")
    d = RP / "definition"
    write(d / "version.json", json.dumps({"$schema": f"{SCHEMA}/versionMetadata/1.0.0/schema.json", "version": "2.0.0"}, indent=2))
    ver = {"visual": "2.13.0", "report": "3.4.0", "page": "2.3.1"}
    write(d / "report.json", json.dumps({
        "$schema": f"{SCHEMA}/report/3.3.0/schema.json",
        "themeCollection": {
            "baseTheme": {"name": base, "reportVersionAtImport": ver, "type": "SharedResources"},
            "customTheme": {"name": "SupplyChainTheme.json", "reportVersionAtImport": ver, "type": "RegisteredResources"},
        },
        "objects": {"section": [{"properties": {"verticalAlignment": {"expr": {"Literal": {"Value": "'Top'"}}}}}]},
        "resourcePackages": [
            {"name": "SharedResources", "type": "SharedResources",
             "items": [{"name": base, "path": f"BaseThemes/{base}.json", "type": "BaseTheme"}]},
            {"name": "RegisteredResources", "type": "RegisteredResources",
             "items": [{"name": "SupplyChainTheme.json", "path": "SupplyChainTheme.json", "type": "CustomTheme"}]},
        ],
        "settings": {"useStylableVisualContainerHeader": True, "exportDataMode": "AllowSummarized",
                     "defaultDrillFilterOtherVisuals": True, "allowChangeFilterTypes": True,
                     "useEnhancedTooltips": True, "useDefaultAggregateDisplayName": True},
    }, indent=2))
    pdir = d / "pages"
    if pdir.exists():
        shutil.rmtree(pdir)
    order = []
    for page_id, display, visuals in pages:
        order.append(page_id)
        write(pdir / page_id / "page.json", json.dumps({
            "$schema": f"{SCHEMA}/page/2.1.0/schema.json", "name": page_id, "displayName": display,
            "displayOption": "FitToPage", "height": 720, "width": 1280}, indent=2))
        for v in visuals:
            write(pdir / page_id / "visuals" / v["name"] / "visual.json", json.dumps(v, indent=2))
    write(pdir / "pages.json", json.dumps({"$schema": f"{SCHEMA}/pagesMetadata/1.1.0/schema.json",
                                           "pageOrder": order, "activePageName": order[0]}, indent=2))


def build_pbip():
    write(PB / f"{NAME}.pbip", json.dumps({"version": "1.0", "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
                                           "settings": {"enableAutoRecovery": True}}, indent=2))
    write(PB / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")


def report_defaults():
    """Default slicer selections: highest-risk supplier and highest-spend product."""
    import pandas as pd
    rep = config.REPORT_DIR / "supply_chain_report.xlsx"
    try:
        sc = pd.read_excel(rep, "Supplier Scorecard")
        abc = pd.read_excel(rep, "ABC-XYZ")
        return abc.iloc[0]["product_name"], sc.iloc[0]["supplier_name"]
    except Exception:
        return None, None


def main(data_folder: str):
    if SM.exists():
        shutil.rmtree(SM)
    build_semantic_model(data_folder)
    import powerbi_visuals
    top_product, top_supplier = report_defaults()
    build_report(powerbi_visuals.pages(top_product, top_supplier))
    build_pbip()
    print(f"Saved powerbi/{NAME}.pbip (data folder: {data_folder})")


if __name__ == "__main__":
    import os
    folder = sys.argv[1] if len(sys.argv) > 1 else str(config.DATA_DIR)
    if not folder.endswith(("\\", "/")):
        folder += os.sep
    main(folder)
