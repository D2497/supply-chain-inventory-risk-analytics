"""
Report pages and visuals for the Power BI project (PBIR format, visualContainer 2.13.0).
Page size 1280 x 720. Used by src/build_powerbi.py.
"""
import hashlib

SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.13.0/schema.json"
M = "_Measures"


# ---------------------------------------------------------------- field helpers
def _src(entity):
    return {"Expression": {"SourceRef": {"Entity": entity}}}


def meas(name, entity=M):
    return ("Measure", entity, name)


def col(entity, name):
    return ("Column", entity, name)


def field(f):
    kind, entity, prop = f
    return {kind: {**_src(entity), "Property": prop}}


def proj(f, display=None):
    kind, entity, prop = f
    p = {"field": field(f), "queryRef": f"{entity}.{prop}", "nativeQueryRef": display or prop}
    if display:
        p["displayName"] = display
    return p


def lit(v):
    if isinstance(v, bool):
        s = "true" if v else "false"
    elif isinstance(v, (int, float)):
        s = f"{v}D"
    else:
        s = "'" + str(v).replace("'", "''") + "'"
    return {"expr": {"Literal": {"Value": s}}}


def color_by_measure(measure_name):
    return {"solid": {"color": {"expr": {"Measure": {**_src(M), "Property": measure_name}}}}}


_counter = {"n": 0}


def vid(page, kind):
    _counter["n"] += 1
    return hashlib.md5(f"{page}/{kind}/{_counter['n']}".encode()).hexdigest()[:20]


def container(page, vtype, x, y, w, h, query=None, objects=None, title=None, sort=None, extra_visual=None):
    v = {"visualType": vtype}
    if query is not None:
        v["query"] = {"queryState": {role: {"projections": [proj(f, d) for f, d in fields]}
                                     for role, fields in query.items()}}
        if sort:
            f, direction = sort
            v["query"]["sortDefinition"] = {"sort": [{"field": field(f), "direction": direction}], "isDefaultSort": True}
    if objects:
        v["objects"] = objects
    vco = {}
    if title:
        vco["title"] = [{"properties": {"show": lit(True), "text": lit(title)}}]
        vco["subTitle"] = [{"properties": {"show": lit(False)}}]
    vco["border"] = [{"properties": {"show": lit(True), "radius": lit(8)}}]
    vco["dropShadow"] = [{"properties": {"show": lit(False)}}]
    v["visualContainerObjects"] = vco
    if extra_visual:
        v.update(extra_visual)
    v["drillFilterOtherVisuals"] = True
    _counter["z"] = _counter.get("z", 0) + 1
    return {"$schema": SCHEMA, "name": vid(page, vtype),
            "position": {"x": x, "y": y, "z": _counter["z"] * 1000, "height": h, "width": w, "tabOrder": _counter["z"] * 1000},
            "visual": v}


# ---------------------------------------------------------------- visual builders
def header(page, title, subtitle):
    paras = [
        {"textRuns": [{"value": title, "textStyle": {"fontWeight": "bold", "fontSize": "20pt", "color": "#123D73"}}]},
        {"textRuns": [{"value": subtitle, "textStyle": {"fontSize": "10pt", "color": "#52514E"}}]},
    ]
    c = container(page, "textbox", 20, 8, 1240, 64, objects={"general": [{"properties": {"paragraphs": paras}}]})
    c["visual"]["visualContainerObjects"]["border"] = [{"properties": {"show": lit(False)}}]
    c["visual"]["visualContainerObjects"]["background"] = [{"properties": {"show": lit(False)}}]
    return c


def card(page, x, y, w, h, measure, label):
    return container(page, "card", x, y, w, h, query={"Values": [(meas(measure), label)]},
                     objects={"categoryLabels": [{"properties": {"show": lit(True)}}],
                              "labels": [{"properties": {"fontSize": lit(22)}}]})


def bar(page, x, y, w, h, category, value, title, color_measure=None, sort_desc=True, cat_label=None, val_label=None):
    objects = {"legend": [{"properties": {"show": lit(False)}}],
               "labels": [{"properties": {"show": lit(True)}}]}
    if color_measure:
        objects["dataPoint"] = [{"properties": {"fill": color_by_measure(color_measure)},
                                 "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}}]
    return container(page, "clusteredBarChart", x, y, w, h,
                     query={"Category": [(category, cat_label)], "Y": [(value, val_label)]},
                     objects=objects, title=title,
                     sort=(value, "Descending") if sort_desc else (category, "Ascending"))


def table(page, x, y, w, h, fields, title, sort=None, font_color=None):
    objects = {}
    if font_color:
        target, color_measure = font_color
        kind, entity, prop = target
        objects["values"] = [{"properties": {"fontColor": color_by_measure(color_measure)},
                              "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}],
                                           "metadata": f"{entity}.{prop}"}}]
    objects["total"] = [{"properties": {"totals": lit(False)}}]
    return container(page, "tableEx", x, y, w, h, query={"Values": fields}, objects=objects,
                     title=title, sort=sort)


def slicer_filter(f, value):
    """Default selection saved with the slicer (same JSON Power BI writes)."""
    kind, entity, prop = f
    src = {"Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": prop}}
    if isinstance(value, str):
        cond = {"In": {"Expressions": [src], "Values": [[{"Literal": {"Value": "'" + value.replace("'", "''") + "'"}}]]}}
    else:
        cond = {"Comparison": {"ComparisonKind": 0, "Left": src, "Right": {"Literal": {"Value": f"{value}D"}}}}
    return [{"properties": {"filter": {"filter": {"Version": 2, "From": [{"Name": "s", "Entity": entity, "Type": 0}],
                                                  "Where": [{"Condition": cond}]}}}}]


def slicer(page, x, y, w, h, f, title, mode="Dropdown", single=False, default=None):
    objects = {"data": [{"properties": {"mode": lit(mode)}}],
               "header": [{"properties": {"show": lit(True), "text": lit(title)}}]}
    if default is not None:
        objects["general"] = slicer_filter(f, default)
        if mode == "Single":   # slider position must be stored separately
            objects["data"][0]["properties"]["numericStart"] = lit(default)
    if single:
        objects["selection"] = [{"properties": {"singleSelect": lit(True), "strictSingleSelect": lit(True)}}]
    return container(page, "slicer", x, y, w, h, query={"Values": [(f, None)]}, objects=objects)


def forecast_settings(measure, length=12, seasonality=52):
    """Power BI's built-in forecast (exponential smoothing), as Power BI saves it."""
    num = lambda v: {"Literal": {"Value": f"{v}D"}}
    return {"properties": {
        "show": lit(True), "displayName": lit("Forecast"),
        "transform": {"algorithm": "Forecast", "parameters": [
            {**num(7), "Name": "Unit"}, {**num(length), "Name": "ForecastLength"},
            {**num(0), "Name": "IgnoreLast"}, {**num(0.95), "Name": "ConfidenceLevel"},
            {**num(seasonality), "Name": "MaxSeasonality"}]}},
        "selector": {"metadata": f"{M}.{measure}", "id": "1"}}


def cards_row(page, y, items, x0=20, total_w=1240, h=100, gap=14):
    w = (total_w - gap * (len(items) - 1)) // len(items)
    return [card(page, x0 + i * (w + gap), y, w, h, m, label) for i, (m, label) in enumerate(items)]


# ---------------------------------------------------------------- pages
def pages(top_product=None, top_risk_supplier=None):
    P = "products"
    S = "suppliers"

    ov = "overview"
    overview = [
        header(ov, "Supply Chain Control Tower",
               "Inventory optimization, demand and supplier risk at a glance. Use the slicers to filter by category or supplier country."),
        *cards_row(ov, 80, [("Annual Spend Total", "Annual spend"), ("A-Class Share of Spend", "Spend in A-class SKUs"),
                            ("SKUs to Order", "SKUs to order now"), ("EOQ Annual Saving", "EOQ saving / year"),
                            ("Excess Inventory Value", "Excess inventory"), ("On-Time Delivery %", "On-time delivery")]),
        bar(ov, 20, 195, 390, 255, col("Status List", "Status"), meas("SKUs in Status"), "SKUs by inventory status",
            color_measure="Status Bar Color", sort_desc=False, cat_label="Status", val_label="SKUs"),
        bar(ov, 424, 195, 470, 255, col(S, "supplier_name"), meas("Risk Score"), "Supplier risk score (0-100)",
            color_measure="Risk Color", cat_label="Supplier", val_label="Risk score"),
        container(ov, "pivotTable", 908, 195, 352, 255,
                  query={"Rows": [(col(P, "ABC"), "ABC")], "Columns": [(col(P, "XYZ"), "XYZ")],
                         "Values": [(meas("SKU Count"), "SKUs")]},
                  title="ABC-XYZ segmentation (SKU count)"),
        slicer(ov, 20, 464, 280, 110, col(P, "category"), "Category", mode="Dropdown"),
        slicer(ov, 20, 588, 280, 110, col(S, "country"), "Supplier country", mode="Dropdown"),
        table(ov, 314, 464, 946, 234,
              [(col(P, "sku"), "SKU"), (col(P, "product_name"), "Product"), (meas("Inventory Status"), "Status"),
               (col(P, "Inventory Position"), "Position"), (meas("Reorder Point"), "Reorder point"),
               (meas("Suggested Order Qty"), "Order qty")],
              "What to order today (largest orders first)", sort=(meas("Suggested Order Qty"), "Descending"),
              font_color=(meas("Inventory Status"), "Status Color")),
    ]

    inv = "inventory"
    inventory = [
        header(inv, "Inventory Policy",
               "Safety stock, reorder point and EOQ for every SKU. Drag the service level slider to see the policy change."),
        slicer(inv, 20, 80, 300, 100, col("Service Level", "Service Level %"), "Service level (%)", mode="Single", default=95),
        *cards_row(inv, 80, [("SKUs Critical", "Critical (below safety stock)"), ("SKUs to Order", "SKUs to order now"),
                             ("SKUs Overstocked", "Overstocked SKUs"), ("Safety Stock Value", "Safety stock value"),
                             ("Excess Inventory Value", "Excess inventory")], x0=334, total_w=926),
        table(inv, 20, 194, 1240, 506,
              [(col(P, "sku"), "SKU"), (col(P, "product_name"), "Product"), (col(P, "Segment"), "Segment"),
               (col(P, "Avg Weekly Demand"), "Avg weekly demand"), (col(P, "LT Mean Wks"), "Lead time (wks)"),
               (meas("Safety Stock"), "Safety stock"), (meas("Reorder Point"), "Reorder point"), (col(P, "EOQ"), "EOQ"),
               (col(P, "Inventory Position"), "Position"), (meas("Inventory Status"), "Status"),
               (meas("Suggested Order Qty"), "Order qty")],
              "Inventory policy by SKU", sort=(col(P, "Annual Spend"), "Descending"),
              font_color=(meas("Inventory Status"), "Status Color")),
    ]

    sp = "suppliers"
    suppliers = [
        header(sp, "Supplier Scorecard",
               "Delivery, lead time, quality and spend concentration, combined into a weighted 0-100 risk score."),
        *cards_row(sp, 80, [("On-Time Delivery %", "On-time delivery"), ("Avg Lead Time Wks", "Avg lead time (wks)"),
                            ("Defect Rate", "Defect rate"), ("Fill Rate", "Fill rate"),
                            ("High-Risk Suppliers", "High-risk suppliers")]),
        table(sp, 20, 194, 1240, 262,
              [(col(S, "supplier_name"), "Supplier"), (col(S, "country"), "Country"), (meas("PO Count"), "POs"),
               (meas("PO Spend"), "Spend"), (meas("Share of Spend"), "Share of spend"),
               (meas("On-Time Delivery %"), "On-time"), (meas("Avg Lead Time Wks"), "Lead time (wks)"),
               (meas("Lead Time CV"), "Lead time CV"), (meas("Defect Rate"), "Defect rate"),
               (meas("Risk Score"), "Risk score"), (meas("Risk Tier"), "Tier"), (meas("Main Risk Driver"), "Main driver")],
              "Supplier scorecard", sort=(meas("Risk Score"), "Descending"),
              font_color=(meas("Risk Tier"), "Risk Color")),
        container(sp, "scatterChart", 20, 470, 640, 230,
                  query={"Category": [(col(S, "supplier_name"), "Supplier")],
                         "X": [(meas("Lead Time CV"), "Lead time CV")],
                         "Y": [(meas("On-Time Delivery %"), "On-time delivery")],
                         "Size": [(meas("PO Spend"), "Spend")]},
                  objects={"categoryLabels": [{"properties": {"show": lit(True)}}]},
                  title="Reliability: on-time delivery vs lead-time variability (size = spend)"),
        bar(sp, 674, 470, 586, 230, col(S, "country"), meas("PO Spend"), "Purchase spend by supplier country",
            cat_label="Country", val_label="Spend"),
    ]

    dm = "demand"
    demand = [
        header(dm, "Demand & Forecast",
               "Weekly demand for all products, or pick one product to see its pattern."),
        slicer(dm, 20, 80, 420, 100, col(P, "product_name"), "Product (leave blank for all)", mode="Dropdown"),
        *cards_row(dm, 80, [("Units Sold", "Units sold (2 years)"), ("Annual Spend Total", "Annual spend"),
                            ("SKU Count", "SKUs")], x0=454, total_w=806, h=100),
        container(dm, "lineChart", 20, 194, 1240, 320,
                  query={"Category": [(col("Dates", "Date"), "Week")], "Y": [(meas("Units Sold"), "Units")]},
                  objects={"categoryAxis": [{"properties": {"axisType": lit("Scalar")}}],
                           "forecast": [forecast_settings("Units Sold")]},
                  title="Weekly units sold"),
        bar(dm, 20, 528, 620, 172, col(P, "category"), meas("Units Sold"), "Units sold by category",
            cat_label="Category", val_label="Units"),
        bar(dm, 654, 528, 606, 172, col(P, "Segment"), meas("SKU Count"), "SKUs by ABC-XYZ segment",
            sort_desc=False, cat_label="Segment", val_label="SKUs"),
    ]

    ds = "disruption"
    disruption = [
        header(ds, "Disruption Simulator",
               "What if a supplier is delayed? Pick a supplier and a delay to see which SKUs stock out and the revenue at risk."),
        slicer(ds, 20, 80, 420, 100, col(S, "supplier_name"), "Delayed supplier", mode="Dropdown", single=True,
               default=top_risk_supplier),
        slicer(ds, 454, 80, 420, 100, col("Delay Weeks", "Delay Weeks"), "Extra delay (weeks)", mode="Single", default=4),
        *cards_row(ds, 194, [("SKUs Stocking Out", "SKUs that stock out"), ("Revenue at Risk", "Revenue at risk"),
                             ("Extra Buffer Cost", "Extra buffer stock to prevent it")]),
        table(ds, 20, 308, 1240, 392,
              [(col(P, "sku"), "SKU"), (col(P, "product_name"), "Product"), (meas("Weeks of Stock"), "Weeks of stock"),
               (col(P, "LT Mean Wks"), "Lead time (wks)"), (meas("New Lead Time Wks"), "New lead time (wks)"),
               (meas("Disruption Outcome"), "Outcome"), (meas("Units Short"), "Units short"),
               (meas("Revenue at Risk"), "Revenue at risk")],
              "Impact by SKU", sort=(meas("Revenue at Risk"), "Descending"),
              font_color=(meas("Disruption Outcome"), "Outcome Color")),
    ]

    lc = "landed"
    LS, LCY = "landed_cost_by_sku", "landed_cost_by_country"
    stacked = container(lc, "barChart", 20, 194, 430, 506,
                        query={"Category": [(col(LCY, "country"), "Country")],
                               "Y": [(meas("Price"), "Product price"), (meas("Freight"), "Freight & insurance"),
                                     (meas("Duty"), "Import duty"), (meas("Brokerage"), "Brokerage & handling")]},
                        objects={"legend": [{"properties": {"show": lit(True), "position": lit("Top")}}],
                                 "labels": [{"properties": {"show": lit(False)}}]},
                        title="Cost to land CAD 100 of product (China price level)",
                        sort=(meas("Price"), "Ascending"))
    landed = [
        header(lc, "Logistics & Landed Cost: Offshore vs Nearshore",
               "Total landed cost = price + freight + duty + brokerage + the cost of safety stock and in-transit inventory a long lead time needs."),
        *cards_row(lc, 80, [("Nearshore Cheaper SKUs", "SKUs cheaper nearshore"),
                            ("Nearshore Saving", "Saving if they move / yr"),
                            ("Nearshore Candidates", "Nearshore candidates"),
                            ("Inventory Released", "Inventory freed"),
                            ("Lead Time Cut Wks", "Avg lead-time cut (wks)")]),
        stacked,
        slicer(lc, 464, 194, 796, 70, col(LS, "recommendation"), "Recommendation", mode="Dropdown"),
        table(lc, 464, 274, 796, 426,
              [(col(LS, "sku"), "SKU"), (col(LS, "product_name"), "Product"),
               (col(LS, "current_country"), "Now"), (col(LS, "nearshore_country"), "Nearshore"),
               (col(LS, "current_lead_time_wks"), "LT now"), (col(LS, "nearshore_lead_time_wks"), "LT near"),
               (meas("Landed Saving"), "Saving/yr"), (col(LS, "inventory_released_cad"), "Inv. freed"),
               (col(LS, "recommendation"), "Recommendation")],
              "Offshore SKUs: total landed cost comparison", sort=(meas("Landed Saving"), "Descending"),
              font_color=(col(LS, "recommendation"), "Recommendation Color")),
    ]

    return [
        ("overview", "Overview", overview),
        ("inventory", "Inventory Policy", inventory),
        ("suppliers", "Supplier Scorecard", suppliers),
        ("demand", "Demand & Forecast", demand),
        ("disruption", "Disruption Simulator", disruption),
        ("landed", "Landed Cost & Nearshoring", landed),
    ]
