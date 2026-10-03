// =====================================================================
// Power Query (M) - load the five CSV files with the right data types.
// Easiest way: Home > Get data > Text/CSV for each file (Power BI detects
// types automatically). If types come out wrong, open Transform data >
// select the query > Advanced Editor, and paste the matching block below.
// Change DataFolder to where YOUR project's data folder is.
// =====================================================================

// ---- 1. Create a parameter first: Home > Manage parameters > New
//      Name: DataFolder   Type: Text   Value: C:\Projects\supply-chain-inventory-risk-analytics\data\

// ---- products
let
    Source = Csv.Document(File.Contents(DataFolder & "products.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Headers, {{"sku", type text}, {"product_name", type text}, {"category", type text}, {"unit_cost_cad", type number}, {"supplier_id", type text}})
in
    Typed

// ---- suppliers
let
    Source = Csv.Document(File.Contents(DataFolder & "suppliers.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Headers, {{"supplier_id", type text}, {"supplier_name", type text}, {"country", type text}, {"promised_lead_time_wks", Int64.Type}})
in
    Typed

// ---- weekly_demand
let
    Source = Csv.Document(File.Contents(DataFolder & "weekly_demand.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Headers, {{"week_start", type date}, {"sku", type text}, {"units", Int64.Type}}, "en-CA")
in
    Typed

// ---- purchase_orders
let
    Source = Csv.Document(File.Contents(DataFolder & "purchase_orders.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Headers, {
        {"po_id", type text}, {"supplier_id", type text}, {"sku", type text},
        {"order_date", type date}, {"promised_date", type date}, {"received_date", type date},
        {"qty_ordered", Int64.Type}, {"qty_received", Int64.Type}, {"qty_defective", Int64.Type},
        {"po_value_cad", type number}}, "en-CA")
in
    Typed

// ---- inventory_snapshot
let
    Source = Csv.Document(File.Contents(DataFolder & "inventory_snapshot.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Headers, {{"sku", type text}, {"on_hand_units", Int64.Type}, {"on_order_units", Int64.Type}})
in
    Typed
