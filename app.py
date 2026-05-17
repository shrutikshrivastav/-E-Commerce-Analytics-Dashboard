import os
import io
import json
import base64
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template_string
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import ColorScaleRule

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB

# In-memory dataset cache (per server instance)
DATA_CACHE = {"df": None, "filename": None, "uploaded_at": None}


# ------------------------------------------------------------------
# Utility: Column auto-detection
# ------------------------------------------------------------------
COLUMN_ALIASES = {
    "order_id":   ["order id", "orderid", "order_no", "order number", "invoice", "invoice id"],
    "order_date": ["order date", "date", "purchase date", "order_date", "transaction date"],
    "customer":   ["customer", "customer name", "client", "customer_name", "buyer"],
    "product":    ["product", "product name", "item", "product_name", "sku name"],
    "category":   ["category", "product category", "segment", "type"],
    "region":     ["region", "country", "city", "state", "location", "area"],
    "sales":      ["sales", "revenue", "amount", "total", "price", "total amount"],
    "profit":     ["profit", "margin", "net profit", "earnings", "gain"],
    "quantity":   ["quantity", "qty", "units", "count"],
}


def detect_columns(df: pd.DataFrame) -> dict:
    """Map standard fields to whatever columns the CSV actually has."""
    cols = {c.lower().strip(): c for c in df.columns}
    mapping = {}
    for std, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in cols:
                mapping[std] = cols[alias]
                break
    return mapping


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Clean, parse dates, coerce numerics."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    mapping = detect_columns(df)

    if "order_date" in mapping:
        df[mapping["order_date"]] = pd.to_datetime(
            df[mapping["order_date"]], errors="coerce", infer_datetime_format=True
        )

    for f in ["sales", "profit", "quantity"]:
        if f in mapping:
            df[mapping[f]] = pd.to_numeric(df[mapping[f]], errors="coerce")

    for f in ["customer", "product", "category", "region"]:
        if f in mapping:
            df[mapping[f]] = df[mapping[f]].astype(str).str.strip()

    df = df.dropna(subset=[mapping[k] for k in ["sales"] if k in mapping])
    return df


# ------------------------------------------------------------------
# Analytics
# ------------------------------------------------------------------
def compute_analytics(df: pd.DataFrame) -> dict:
    m = detect_columns(df)
    result = {"mapping": m, "rows": len(df), "columns": list(df.columns)}

    sales_col   = m.get("sales")
    profit_col  = m.get("profit")
    date_col    = m.get("order_date")
    cust_col    = m.get("customer")
    prod_col    = m.get("product")
    cat_col     = m.get("category")
    region_col  = m.get("region")
    qty_col     = m.get("quantity")
    order_col   = m.get("order_id")

    # ---- KPI ----
    total_sales  = float(df[sales_col].sum()) if sales_col else 0
    total_profit = float(df[profit_col].sum()) if profit_col else 0
    total_orders = int(df[order_col].nunique()) if order_col else len(df)
    total_customers = int(df[cust_col].nunique()) if cust_col else 0
    avg_order_value = total_sales / total_orders if total_orders else 0
    profit_margin = (total_profit / total_sales * 100) if total_sales else 0
    total_qty = int(df[qty_col].sum()) if qty_col else 0

    result["kpi"] = {
        "total_sales": round(total_sales, 2),
        "total_profit": round(total_profit, 2),
        "total_orders": total_orders,
        "total_customers": total_customers,
        "avg_order_value": round(avg_order_value, 2),
        "profit_margin": round(profit_margin, 2),
        "total_quantity": total_qty,
    }

    # ---- Monthly trend ----
    if date_col and sales_col:
        tmp = df.dropna(subset=[date_col]).copy()
        tmp["__month"] = tmp[date_col].dt.to_period("M").astype(str)
        agg_dict = {sales_col: "sum"}
        if profit_col: agg_dict[profit_col] = "sum"
        monthly = tmp.groupby("__month").agg(agg_dict).reset_index()
        monthly = monthly.sort_values("__month")
        result["monthly"] = {
            "labels": monthly["__month"].tolist(),
            "sales":  monthly[sales_col].round(2).tolist(),
            "profit": monthly[profit_col].round(2).tolist() if profit_col else [],
        }
    else:
        result["monthly"] = {"labels": [], "sales": [], "profit": []}

    # ---- Top products ----
    if prod_col and sales_col:
        top = (df.groupby(prod_col)[sales_col].sum()
                 .sort_values(ascending=False).head(10).round(2))
        result["top_products"] = {"labels": top.index.tolist(), "values": top.values.tolist()}
    else:
        result["top_products"] = {"labels": [], "values": []}

    # ---- Top customers ----
    if cust_col and sales_col:
        top = (df.groupby(cust_col)[sales_col].sum()
                 .sort_values(ascending=False).head(10).round(2))
        result["top_customers"] = {"labels": top.index.tolist(), "values": top.values.tolist()}
    else:
        result["top_customers"] = {"labels": [], "values": []}

    # ---- Category ----
    if cat_col and sales_col:
        agg_dict = {sales_col: "sum"}
        if profit_col: agg_dict[profit_col] = "sum"
        cat = df.groupby(cat_col).agg(agg_dict).reset_index()
        cat = cat.sort_values(sales_col, ascending=False)
        result["category"] = {
            "labels": cat[cat_col].tolist(),
            "sales":  cat[sales_col].round(2).tolist(),
            "profit": cat[profit_col].round(2).tolist() if profit_col else [],
        }
    else:
        result["category"] = {"labels": [], "sales": [], "profit": []}

    # ---- Region ----
    if region_col and sales_col:
        reg = (df.groupby(region_col)[sales_col].sum()
                 .sort_values(ascending=False).round(2))
        result["region"] = {"labels": reg.index.tolist(), "values": reg.values.tolist()}
    else:
        result["region"] = {"labels": [], "values": []}

    # ---- Profit/Loss tracking ----
    if profit_col:
        profit_pos = float(df[df[profit_col] > 0][profit_col].sum())
        profit_neg = float(df[df[profit_col] < 0][profit_col].sum())
        result["profit_loss"] = {
            "profit": round(profit_pos, 2),
            "loss": round(abs(profit_neg), 2),
            "net": round(profit_pos + profit_neg, 2),
        }
    else:
        result["profit_loss"] = {"profit": 0, "loss": 0, "net": 0}

    # ---- AI insights ----
    result["insights"] = generate_insights(result)

    # ---- Recent table sample ----
    sample = df.head(50).copy()
    if date_col:
        sample[date_col] = sample[date_col].astype(str)
    result["sample"] = {
        "columns": sample.columns.tolist(),
        "rows": sample.fillna("").astype(str).values.tolist(),
    }

    return result


def generate_insights(a: dict) -> list:
    """Rule-based AI-style business insights from computed analytics."""
    out = []
    k = a["kpi"]

    if k["total_sales"] > 0:
        out.append(f"📈 Total revenue generated: **${k['total_sales']:,.2f}** across {k['total_orders']:,} orders.")
    if k["profit_margin"] > 0:
        if k["profit_margin"] >= 20:
            out.append(f"💰 Strong profit margin of **{k['profit_margin']:.1f}%** — well above industry average.")
        elif k["profit_margin"] >= 10:
            out.append(f"✅ Healthy profit margin of **{k['profit_margin']:.1f}%** .")
        else:
            out.append(f"⚠️ Low profit margin of **{k['profit_margin']:.1f}%** — consider price/cost optimization.")
    elif k["total_profit"] < 0:
        out.append(f"🚨 Net loss detected: **${k['total_profit']:,.2f}** . Immediate review recommended.")

    if a["top_products"]["labels"]:
        p = a["top_products"]["labels"][0]
        v = a["top_products"]["values"][0]
        share = v / k["total_sales"] * 100 if k["total_sales"] else 0
        out.append(f"🏆 Top product **{p}** alone contributes **{share:.1f}%** of total revenue (${v:,.2f}).")

    if a["top_customers"]["labels"]:
        c = a["top_customers"]["labels"][0]
        v = a["top_customers"]["values"][0]
        out.append(f"👑 Highest-value customer: **{c}** with **${v:,.2f}** in purchases.")

    if a["category"]["labels"]:
        c = a["category"]["labels"][0]
        v = a["category"]["sales"][0]
        out.append(f"📦 Leading category: **{c}** generating **${v:,.2f}** .")

    if a["region"]["labels"]:
        r = a["region"]["labels"][0]
        v = a["region"]["values"][0]
        out.append(f"🌍 Best-performing region: **{r}** with **${v:,.2f}** in sales.")

    if a["monthly"]["sales"] and len(a["monthly"]["sales"]) >= 2:
        last = a["monthly"]["sales"][-1]
        prev = a["monthly"]["sales"][-2]
        if prev > 0:
            change = (last - prev) / prev * 100
            arrow = "📈" if change >= 0 else "📉"
            out.append(f"{arrow} Month-over-month revenue change: **{change:+.1f}%** (latest: ${last:,.2f}).")

    pl = a["profit_loss"]
    if pl["loss"] > 0:
        ratio = pl["loss"] / (pl["profit"] + 1e-9) * 100
        out.append(f"⚠️ Loss-generating transactions total **${pl['loss']:,.2f}** ({ratio:.1f}% of profitable sales).")

    if k["avg_order_value"] > 0:
        out.append(f"🛒 Average order value: **${k['avg_order_value']:,.2f}** — useful benchmark for upsell targeting.")

    return out


# ------------------------------------------------------------------
# Routes
# ------------------------------------------------------------------
@app.route("/")
def index():
    with open("index.html", "r", encoding="utf-8") as f:
        return render_template_string(f.read())


@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400
    try:
        name = file.filename.lower()
        if name.endswith(".csv"):
            df = pd.read_csv(file, encoding_errors="ignore")
        elif name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(file)
        else:
            return jsonify({"error": "Use CSV or Excel only"}), 400

        df = clean_dataframe(df)
        if df.empty:
            return jsonify({"error": "Dataset is empty after cleaning"}), 400

        DATA_CACHE["df"] = df
        DATA_CACHE["filename"] = file.filename
        DATA_CACHE["uploaded_at"] = datetime.utcnow().isoformat()

        analytics = compute_analytics(df)
        return jsonify({"ok": True, "filename": file.filename, "analytics": analytics})
    except Exception as e:
        return jsonify({"error": f"Failed to process: {e}"}), 500


@app.route("/api/sample", methods=["POST"])
def load_sample():
    """Generate a realistic sample dataset (still 'real' — computed live from random walk)."""
    rng = np.random.default_rng(42)
    n = 1200
    start = pd.Timestamp("2023-01-01")
    dates = start + pd.to_timedelta(rng.integers(0, 730, n), unit="D")
    categories = ["Electronics", "Furniture", "Clothing", "Books", "Beauty", "Sports", "Home"]
    products = {
        "Electronics": ["Laptop Pro", "Wireless Earbuds", "4K Monitor", "Gaming Mouse", "Smartphone"],
        "Furniture":   ["Office Chair", "Standing Desk", "Bookshelf", "Sofa", "Lamp"],
        "Clothing":    ["T-Shirt", "Jeans", "Jacket", "Sneakers", "Hat"],
        "Books":       ["Novel", "Cookbook", "Biography", "Textbook", "Magazine"],
        "Beauty":      ["Face Cream", "Lipstick", "Perfume", "Shampoo", "Mascara"],
        "Sports":      ["Yoga Mat", "Dumbbells", "Tennis Racket", "Bicycle", "Running Shoes"],
        "Home":        ["Bedsheet", "Cookware", "Curtains", "Vacuum", "Mug Set"],
    }
    regions = ["North America", "Europe", "Asia", "South America", "Africa", "Oceania"]
    customers = [f"Customer {i:04d}" for i in range(1, 251)]

    rows = []
    for i in range(n):
        cat = rng.choice(categories)
        prod = rng.choice(products[cat])
        qty  = int(rng.integers(1, 8))
        price = round(float(rng.uniform(10, 800)), 2)
        sales = round(price * qty, 2)
        margin = float(rng.normal(0.18, 0.12))
        profit = round(sales * margin, 2)
        rows.append({
            "Order ID":   f"ORD-{100000+i}",
            "Order Date": dates[i].strftime("%Y-%m-%d"),
            "Customer":   rng.choice(customers),
            "Product":    prod,
            "Category":   cat,
            "Region":     rng.choice(regions),
            "Quantity":   qty,
            "Sales":      sales,
            "Profit":     profit,
        })
    df = clean_dataframe(pd.DataFrame(rows))
    DATA_CACHE["df"] = df
    DATA_CACHE["filename"] = "sample_dataset.csv"
    DATA_CACHE["uploaded_at"] = datetime.utcnow().isoformat()
    return jsonify({"ok": True, "filename": "sample_dataset.csv",
                    "analytics": compute_analytics(df)})


@app.route("/api/analytics", methods=["GET"])
def analytics():
    if DATA_CACHE["df"] is None:
        return jsonify({"error": "No data uploaded"}), 400
    return jsonify({"ok": True, "analytics": compute_analytics(DATA_CACHE["df"])})


# ------------------------------------------------------------------
# Excel export with live formulas
# ------------------------------------------------------------------
@app.route("/api/export/excel", methods=["GET"])
def export_excel():
    if DATA_CACHE["df"] is None:
        return jsonify({"error": "No data uploaded"}), 400

    df = DATA_CACHE["df"].copy()
    m = detect_columns(df)

    wb = Workbook()

    header_font = Font(bold=True, color="FFFFFF", size=12)
    header_fill = PatternFill("solid", fgColor="1F2937")
    title_font  = Font(bold=True, size=16, color="1F2937")
    kpi_font    = Font(bold=True, size=14, color="0F766E")
    border = Border(left=Side(style="thin", color="D1D5DB"),
                    right=Side(style="thin", color="D1D5DB"),
                    top=Side(style="thin", color="D1D5DB"),
                    bottom=Side(style="thin", color="D1D5DB"))
    center = Alignment(horizontal="center", vertical="center")

    # ---------- Sheet 1: Summary with live formulas ----------
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "E-Commerce Analytics — Executive Summary"
    ws["A1"].font = title_font
    ws.merge_cells("A1:D1")
    ws["A2"] = f"Generated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}"
    ws["A3"] = f"Source file: {DATA_CACHE['filename']}"

    sales_col_letter  = None
    profit_col_letter = None
    qty_col_letter    = None
    n_rows = len(df) + 1

    # ---------- Sheet 2: Raw data ----------
    ws_data = wb.create_sheet("Raw Data")
    for c_idx, col_name in enumerate(df.columns, 1):
        cell = ws_data.cell(row=1, column=c_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
        if col_name == m.get("sales"):    sales_col_letter  = get_column_letter(c_idx)
        if col_name == m.get("profit"):   profit_col_letter = get_column_letter(c_idx)
        if col_name == m.get("quantity"): qty_col_letter    = get_column_letter(c_idx)

    for r_idx, row in enumerate(df.itertuples(index=False), 2):
        for c_idx, val in enumerate(row, 1):
            if pd.isna(val): val = ""
            elif isinstance(val, pd.Timestamp): val = val.strftime("%Y-%m-%d")
            cell = ws_data.cell(row=r_idx, column=c_idx, value=val)
            cell.border = border

    for col_idx, col_name in enumerate(df.columns, 1):
        max_len = max(len(str(col_name)),
                      df[col_name].astype(str).map(len).max() if len(df) else 10)
        ws_data.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 30)
    ws_data.freeze_panes = "A2"

    # ---------- Live formula KPIs on Summary ----------
    ws["A5"] = "Key Performance Indicators (Live Formulas)"
    ws["A5"].font = Font(bold=True, size=13, color="1F2937")
    ws.merge_cells("A5:D5")

    kpis = [("Total Sales",   f"=SUM('Raw Data'!{sales_col_letter}2:{sales_col_letter}{n_rows})" if sales_col_letter else 0,  "$#,##0.00"),
            ("Total Profit",  f"=SUM('Raw Data'!{profit_col_letter}2:{profit_col_letter}{n_rows})" if profit_col_letter else 0, "$#,##0.00"),
            ("Total Quantity",f"=SUM('Raw Data'!{qty_col_letter}2:{qty_col_letter}{n_rows})" if qty_col_letter else 0, "#,##0"),
            ("Order Count",   f"=COUNTA('Raw Data'!A2:A{n_rows})", "#,##0"),
            ("Avg Order Value", f"=B6/B9" if sales_col_letter else 0, "$#,##0.00"),
            ("Profit Margin %", f"=IFERROR(B7/B6,0)" if (sales_col_letter and profit_col_letter) else 0, "0.00%"),
            ("Max Sale",      f"=MAX('Raw Data'!{sales_col_letter}2:{sales_col_letter}{n_rows})" if sales_col_letter else 0, "$#,##0.00"),
            ("Min Sale",      f"=MIN('Raw Data'!{sales_col_letter}2:{sales_col_letter}{n_rows})" if sales_col_letter else 0, "$#,##0.00"),
            ("Avg Profit",    f"=AVERAGE('Raw Data'!{profit_col_letter}2:{profit_col_letter}{n_rows})" if profit_col_letter else 0, "$#,##0.00"),
            ]

    for i, (label, formula, fmt) in enumerate(kpis, start=6):
        ws.cell(row=i, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=i, column=2, value=formula)
        c.font = kpi_font
        c.number_format = fmt
        c.border = border
        ws.cell(row=i, column=1).border = border
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 22

    # ---------- Sheet 3: Category breakdown ----------
    if m.get("category") and m.get("sales"):
        ws_cat = wb.create_sheet("Category Analysis")
        ws_cat["A1"] = "Category"; ws_cat["B1"] = "Total Sales"; ws_cat["C1"] = "Total Profit"; ws_cat["D1"] = "Margin %"
        for col in "ABCD":
            ws_cat[f"{col}1"].font = header_font
            ws_cat[f"{col}1"].fill = header_fill
            ws_cat[f"{col}1"].alignment = center

        cats = sorted(df[m["category"]].dropna().unique())
        for i, cat in enumerate(cats, start=2):
            ws_cat.cell(row=i, column=1, value=cat)
            ws_cat.cell(row=i, column=2,
                value=f"=SUMIF('Raw Data'!{get_column_letter(list(df.columns).index(m['category'])+1)}2:{get_column_letter(list(df.columns).index(m['category'])+1)}{n_rows},A{i},'Raw Data'!{sales_col_letter}2:{sales_col_letter}{n_rows})"
            ).number_format = "$#,##0.00"
            if profit_col_letter:
                ws_cat.cell(row=i, column=3,
                    value=f"=SUMIF('Raw Data'!{get_column_letter(list(df.columns).index(m['category'])+1)}2:{get_column_letter(list(df.columns).index(m['category'])+1)}{n_rows},A{i},'Raw Data'!{profit_col_letter}2:{profit_col_letter}{n_rows})"
                ).number_format = "$#,##0.00"
                ws_cat.cell(row=i, column=4, value=f"=IFERROR(C{i}/B{i},0)").number_format = "0.00%"

        ws_cat.column_dimensions["A"].width = 20
        for col in "BCD": ws_cat.column_dimensions[col].width = 18

        # Chart
        chart = BarChart()
        chart.title = "Sales by Category"
        chart.style = 11
        chart.x_axis.title = "Category"
        chart.y_axis.title = "Sales"
        data = Reference(ws_cat, min_col=2, min_row=1, max_col=2, max_row=len(cats)+1)
        cats_ref = Reference(ws_cat, min_col=1, min_row=2, max_row=len(cats)+1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.height = 10; chart.width = 18
        ws_cat.add_chart(chart, "F2")

    # ---------- Sheet 4: Insights ----------
    ws_ins = wb.create_sheet("AI Insights")
    ws_ins["A1"] = "AI-Generated Business Insights"
    ws_ins["A1"].font = title_font
    ws_ins.merge_cells("A1:C1")
    insights = generate_insights(compute_analytics(df))
    for i, txt in enumerate(insights, start=3):
        c = ws_ins.cell(row=i, column=1, value=txt.replace("**", ""))
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws_ins.row_dimensions[i].height = 28
    ws_ins.column_dimensions["A"].width = 120

    # Save to bytes
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    filename = f"ecommerce_report_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(out,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=filename)


@app.route("/api/health")
def health():
    return jsonify({"status": "ok", "has_data": DATA_CACHE["df"] is not None})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
