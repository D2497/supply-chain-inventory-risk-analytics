# Power BI report: Supply Chain Control Tower

A 6-page Power BI report built on a star-schema data model with 50+ DAX measures and two what-if parameters (service level and supplier delay). It's saved as a **Power BI Project (`.pbip`)**, so the data model (TMDL) and the report (PBIR JSON) are plain text files you can read and compare on GitHub.

![Overview page](screenshots/01_overview.png)

## Open the finished report (2 minutes)

1. Install **Power BI Desktop** (free, Windows only) from the Microsoft Store.
2. Run `python src/main.py` once, so the `data/` folder has the CSV files (including `landed_cost_by_sku.csv` and `landed_cost_by_country.csv`, which Python calculates and Power BI displays).
3. Double-click **`powerbi/SupplyChainControlTower.pbix`** (single file), or **`SupplyChainControlTower.pbip`** (the same report as readable project files).
4. If Power BI can't find the data, go to **Home > Transform data > Edit parameters** and set **DataFolder** to your project's `data` folder, ending with a backslash, for example `C:\Users\you\Documents\supply-chain-inventory-risk-analytics\data\`. Then click **Refresh**.
   You can also run `python src/build_powerbi.py`, which regenerates the `.pbip` project with the right path for your computer. The regenerated copy doesn't include the forecast line (see the tip below), so add it again if you use it.

| Page | What it shows |
|---|---|
| Overview | KPI cards, SKUs by inventory status, supplier risk ranking, ABC-XYZ matrix, "what to order today" |
| Inventory Policy | Safety stock, reorder point, EOQ and status per SKU, with a **service level slider** (90–99%) |
| Supplier Scorecard | On-time delivery, lead time, defects, spend share, risk score, tier and main risk driver |
| Demand & Forecast | Weekly demand with a 12-week seasonal forecast (Power BI's exponential smoothing), units by category, SKUs by segment |
| Disruption Simulator | Pick a supplier and a **delay slider** to see which SKUs stock out and the revenue at risk |
| Landed Cost & Nearshoring | Cost to land CAD 100 by country (price, freight, duty, brokerage), and offshore vs nearshore total cost per SKU with a recommendation |

How it was built: `src/build_powerbi.py` generates the semantic model (tables, Power Query, relationships, calculated columns, measures), and `src/powerbi_visuals.py` defines every page and visual. The numbers were checked against the Excel model: excess inventory, safety stock value, EOQ saving and the disruption results match exactly.

**Tip:** the forecast on the Demand page comes from the **Analytics pane > Forecast** (12 points, seasonality 52, 95% confidence band).

| | |
|---|---|
| ![Inventory policy](screenshots/02_inventory_policy.png) | ![Supplier scorecard](screenshots/03_supplier_scorecard.png) |
| ![Demand and forecast](screenshots/04_demand_forecast.png) | ![Disruption simulator](screenshots/05_disruption_simulator.png) |
| ![Landed cost and nearshoring](screenshots/06_landed_cost_nearshoring.png) | |

---

# Build it yourself from scratch (learning guide)

Want to learn Power BI by building the same report by hand? Follow these steps (2–3 hours). Interviewers often ask how you built the model, so this is worth doing at least once.

| File | What it is |
|---|---|
| `measures.dax` | Every calculated column, table and measure, ready to paste |
| `power_query.m` | Power Query code that loads the CSVs with correct data types |
| `SupplyChainTheme.json` | Colour theme (colourblind-checked palette) |

## Step 1: Load the data (10 min)

1. Run `python src/main.py` once, so the `data/` folder has fresh CSVs (simulated or Kaggle).
2. Open Power BI Desktop. Go to **Home > Get data > Text/CSV** and load these 5 files from `data/`, one at a time:
   `products.csv`, `suppliers.csv`, `weekly_demand.csv`, `purchase_orders.csv`, `inventory_snapshot.csv`
   Click **Load** for each one.
3. Click **Transform data** and check the column types: the dates should show a calendar icon and the numbers should show 1.2 or 123. If something is wrong, use the code in `power_query.m`.
4. Go to **View > Themes > Browse for themes** and pick `SupplyChainTheme.json`.

## Step 2: Build the data model (10 min)

Open **Model view** (the third icon on the left) and drag fields to create these relationships. All are many-to-one, single direction:

```
weekly_demand[sku]        ──►  products[sku]
products[supplier_id]     ──►  suppliers[supplier_id]
purchase_orders[supplier_id] ──►  suppliers[supplier_id]
weekly_demand[week_start] ──►  Calendar[Date]          (after Step 3 creates Calendar)
```

Power BI may create extra relationships automatically. **Delete any between `purchase_orders[sku]` and `products[sku]`, or between `inventory_snapshot` and `products`.** The DAX looks those values up itself, and the extra paths would make the model ambiguous.

> Interview point: this is a **star schema**. `products` and `suppliers` are dimension tables, and `weekly_demand` and `purchase_orders` are fact tables.

## Step 3: Add calculated columns, tables and parameters (30 min)

Open `measures.dax` and work through it in order:

1. **Section 1**, on `purchase_orders`: select the table in Data view, go to **Table tools > New column**, and paste each formula (one column at a time).
2. **Section 2**, on `suppliers`: the `Country Risk` column.
3. **Section 3**, on `products`: all the columns from `Avg Weekly Demand` to `EOQ`, **in the order listed**, because later columns use earlier ones.
4. **Section 4**, calculated tables: go to **Modeling > New table** and paste `Calendar`, `Status List` and `_Measures`. Then mark `Calendar` as a date table (**Table tools > Mark as date table**, column `Date`).
5. **What-if parameters**: go to **Modeling > New parameter > Numeric range** and create these two:
   - Name `Service Level`, minimum 0.90, maximum 0.99, increment 0.01, default 0.95. Tick "Add slicer to this page".
   - Name `Delay Weeks`, minimum 0, maximum 12, increment 1, default 4.
6. **Sections 5–7**, measures: select `_Measures`, go to **Home > New measure**, and paste each measure. Then set the formats in **Measure tools**:
   - Percent: `On-Time Delivery %`, `Share of Spend`, `Defect Rate`, `Fill Rate`, `Lead Time CV`, `A-Class Share of Spend`
   - Currency, 0 decimals: `Excess Inventory Value`, `Safety Stock Value`, `Annual Spend Total`, `EOQ Annual Saving`, `Revenue at Risk`, `Extra Buffer Cost`, `PO Spend`
   - Whole number: `Safety Stock`, `Reorder Point`, `Suggested Order Qty`, `Units Short`
7. Go to **Data view > Status List**, select the `Status` column, and set **Column tools > Sort by column > Sort**.

**Check:** add a Table visual with `products[sku]`, `[Safety Stock]`, `[Reorder Point]`, `products[EOQ]` and `[Inventory Status]`. Every row should have values. If you see errors, check the table and column names exactly match your CSV headers.

## Step 4: Build the 5 report pages (1–2 hours)

Use a 16:9 page. On each page, add a text box at the top with a title and a one-line subtitle.

### Page 1: Overview
| Visual | Fields |
|---|---|
| 6 **Card** visuals in a row | `Annual Spend Total`, `A-Class Share of Spend`, `SKUs to Order`, `EOQ Annual Saving`, `Excess Inventory Value`, `On-Time Delivery %` |
| **Clustered bar chart**: "SKUs by inventory status" | Y-axis `Status List[Status]`, X-axis `SKUs in Status`. Bar colours: **Format > Bars > fx > Field value > `Status Color`** |
| **Clustered bar chart**: "Supplier risk score" | Y-axis `suppliers[supplier_name]`, X-axis `Risk Score`, sorted descending, coloured by `Risk Color` |
| **Matrix**: "ABC-XYZ segmentation" | Rows `products[ABC]`, Columns `products[XYZ]`, Values Count of `products[sku]`. Add background colour: **Cell elements > Background color > Gradient** |
| **Slicers** | `products[category]`, `suppliers[country]` |

### Page 2: Inventory policy
| Visual | Fields |
|---|---|
| **Slicer** (from the What-if step) | `Service Level` |
| **Table** | `products[sku]`, `product_name`, `Segment`, `Avg Weekly Demand`, `LT Mean Wks`, `Safety Stock`, `Reorder Point`, `EOQ`, `Inventory Position`, `Inventory Status`, `Suggested Order Qty` |
| Conditional formatting | `Inventory Status` column: **Background color > fx > Field value > `Status Color`** |
| **Cards** | `SKUs Critical`, `SKUs Overstocked`, `Safety Stock Value` |

Drag the Service Level slider from 90% to 99% in your demo. Safety stock, reorder points and statuses all update live. Interviewers enjoy seeing this.

### Page 3: Supplier scorecard
| Visual | Fields |
|---|---|
| **Table** | `supplier_name`, `country`, `PO Count`, `PO Spend`, `Share of Spend`, `On-Time Delivery %`, `Avg Lead Time Wks`, `Lead Time CV`, `Defect Rate`, `Risk Score`, `Risk Tier`, `Main Risk Driver` |
| **Scatter chart** | X `Lead Time CV`, Y `On-Time Delivery %`, Size `PO Spend`, Values `supplier_name` |
| **Filled map** or **Map** | Location `suppliers[country]`, Bubble size `PO Spend`, Tooltips `Risk Score` |
| **Line chart**: "On-time delivery trend" | X `purchase_orders[promised_date]` (Month), Y `On-Time Delivery %` |

### Page 4: Demand & forecast
| Visual | Fields |
|---|---|
| **Slicer** | `products[product_name]` (single select) |
| **Line chart** | X `Calendar[Date]` (set to continuous), Y Sum of `weekly_demand[units]` |
| Built-in forecast | Select the line chart, open the **Analytics pane (magnifying glass) > Forecast > Add**: forecast length 12 points, confidence 95%, seasonality 52 |
| **Cards** | `Avg Weekly Demand` (use Average), `Demand CV` (Average) |

> Interview point: Power BI's forecast uses **exponential smoothing (ETS)**. The Python version compares 5 methods and tests them on weeks they never saw.

### Page 5: Disruption simulator
| Visual | Fields |
|---|---|
| **Slicer** | `suppliers[supplier_name]` (single select) |
| **Slicer** | `Delay Weeks` |
| **Cards** | `SKUs Stocking Out`, `Revenue at Risk`, `Extra Buffer Cost` |
| **Table** | `products[sku]`, `product_name`, `Weeks of Stock`, `New Lead Time Wks`, `Disruption Outcome`, `Units Short`, `Revenue at Risk` |

## Step 5: Finish and publish (20 min)

1. Hide helper tables from report view: right-click `Status List` and `_Measures`, then **Hide in report view** (keep the measures visible).
2. Save as `powerbi/Supply_Chain_Control_Tower.pbix`.
3. Take a screenshot of every page (**Windows + Shift + S**) and save them as `powerbi/screenshots/01_overview.png`, `02_inventory.png` and so on. Add them to the main README.
4. Optional: **File > Export > Export to PDF** and save it as `powerbi/Supply_Chain_Control_Tower.pdf`. Recruiters can open a PDF directly on GitHub.
5. Optional: with a free Power BI account you can **Publish** to the Power BI Service. **Don't use "Publish to web"** unless you're comfortable with the report being fully public.

## Why the numbers can differ from the Python version

- Safety stock in Power BI and Excel uses the **standard deviation of weekly demand** over the last 52 weeks. Python uses **forecast error**, which is a more advanced method. So the statuses can differ slightly between versions, which is a good talking point.
- Supplier on-time delivery, defect rate, lead times and spend share should match Python exactly.
