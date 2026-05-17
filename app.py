import os
import io
import json
import base64
import traceback
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference

app = Flask(__name__, static_folder=".", template_folder=".")
CORS(app)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ─── Helpers ──────────────────────────────────────────────────────────────────

def fig_to_json(fig):
    return json.loads(fig.to_json())


def safe_col(df, *candidates):
    """Return first matching column name (case-insensitive)."""
    lower = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c.lower() in lower:
            return lower[c.lower()]
    return None


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = df.columns.str.strip()
    df = df.dropna(how="all").reset_index(drop=True)

    # Normalise obvious column synonyms
    renames = {}
    mapping = {
        "order_id": ["orderid", "order id", "id", "transaction_id"],
        "date": ["order_date", "orderdate", "sale_date", "purchase_date", "invoice_date"],
        "sales": ["revenue", "amount", "total", "sale_amount", "total_sales", "price"],
        "profit": ["profit_amount", "net_profit", "margin"],
        "quantity": ["qty", "units", "quantity_ordered"],
        "category": ["product_category", "cat", "segment"],
        "sub_category": ["sub_cat", "subcategory", "sub category"],
        "product_name": ["product", "item", "item_name", "product name"],
        "customer_name": ["customer", "client", "buyer", "customer name"],
        "region": ["area", "zone", "territory", "state", "country", "location"],
        "city": ["town", "city_name"],
        "ship_mode": ["shipping_mode", "delivery_mode", "ship mode"],
    }
    lower_cols = {c.lower(): c for c in df.columns}
    for canonical, synonyms in mapping.items():
        if canonical not in lower_cols:
            for syn in synonyms:
                if syn in lower_cols:
                    renames[lower_cols[syn]] = canonical
                    break
    if renames:
        df = df.rename(columns=renames)

    # Parse date
    date_col = safe_col(df, "date", "order_date")
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], infer_datetime_format=True, errors="coerce")

    # Numeric coercion
    for col in ["sales", "profit", "quantity", "discount"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


# ─── Analysis Engine ──────────────────────────────────────────────────────────

def analyse(df: pd.DataFrame) -> dict:
    result = {}

    date_col = safe_col(df, "date", "order_date")
    sales_col = safe_col(df, "sales", "revenue", "amount")
    profit_col = safe_col(df, "profit")
    qty_col = safe_col(df, "quantity", "qty")
    cat_col = safe_col(df, "category", "product_category")
    subcat_col = safe_col(df, "sub_category", "subcategory")
    product_col = safe_col(df, "product_name", "product")
    customer_col = safe_col(df, "customer_name", "customer")
    region_col = safe_col(df, "region", "state", "country")

    # ── KPIs ────────────────────────────────────────────────────────────────
    total_sales = float(df[sales_col].sum()) if sales_col else 0
    total_profit = float(df[profit_col].sum()) if profit_col else 0
    total_orders = int(df.shape[0])
    total_qty = int(df[qty_col].sum()) if qty_col else 0
    avg_order_value = total_sales / total_orders if total_orders else 0
    profit_margin = (total_profit / total_sales * 100) if total_sales else 0
    unique_customers = int(df[customer_col].nunique()) if customer_col else 0

    result["kpis"] = {
        "total_sales": round(total_sales, 2),
        "total_profit": round(total_profit, 2),
        "total_orders": total_orders,
        "total_quantity": total_qty,
        "avg_order_value": round(avg_order_value, 2),
        "profit_margin": round(profit_margin, 2),
        "unique_customers": unique_customers,
    }

    # ── Monthly Revenue Trend ───────────────────────────────────────────────
    if date_col and sales_col:
        monthly = (
            df.dropna(subset=[date_col])
            .groupby(df[date_col].dt.to_period("M"))
            .agg(sales=(sales_col, "sum"), profit=(profit_col, "sum") if profit_col else (sales_col, "count"))
            .reset_index()
        )
        monthly[date_col] = monthly[date_col].astype(str)
        fig_trend = go.Figure()
        fig_trend.add_trace(go.Scatter(
            x=monthly[date_col], y=monthly["sales"],
            name="Revenue", mode="lines+markers",
            line=dict(color="#6ee7b7", width=2.5),
            marker=dict(size=6, color="#6ee7b7"),
            fill="tozeroy", fillcolor="rgba(110,231,183,0.08)"
        ))
        if profit_col:
            fig_trend.add_trace(go.Scatter(
                x=monthly[date_col], y=monthly["profit"],
                name="Profit", mode="lines+markers",
                line=dict(color="#f472b6", width=2, dash="dot"),
                marker=dict(size=5, color="#f472b6")
            ))
        fig_trend.update_layout(**_dark_layout("Monthly Revenue & Profit Trend"))
        result["monthly_trend"] = fig_to_json(fig_trend)

    # ── Category Sales (Donut) ──────────────────────────────────────────────
    if cat_col and sales_col:
        cat_df = df.groupby(cat_col)[sales_col].sum().reset_index().sort_values(sales_col, ascending=False)
        fig_cat = go.Figure(go.Pie(
            labels=cat_df[cat_col], values=cat_df[sales_col],
            hole=0.55,
            marker=dict(colors=["#6ee7b7","#f472b6","#60a5fa","#fbbf24","#a78bfa","#fb923c"]),
            textfont=dict(size=12)
        ))
        fig_cat.update_layout(**_dark_layout("Sales by Category"))
        result["category_chart"] = fig_to_json(fig_cat)

    # ── Sub-Category Bar ────────────────────────────────────────────────────
    if subcat_col and sales_col:
        sub_df = df.groupby(subcat_col)[sales_col].sum().reset_index().sort_values(sales_col, ascending=False).head(12)
        fig_sub = go.Figure(go.Bar(
            x=sub_df[sales_col], y=sub_df[subcat_col],
            orientation="h",
            marker=dict(
                color=sub_df[sales_col],
                colorscale=[[0,"#1e3a5f"],[0.5,"#3b82f6"],[1,"#6ee7b7"]],
                showscale=False
            ),
            text=sub_df[sales_col].apply(lambda v: f"${v:,.0f}"),
            textposition="outside"
        ))
        fig_sub.update_layout(**_dark_layout("Revenue by Sub-Category"))
        result["subcategory_chart"] = fig_to_json(fig_sub)

    # ── Top Products ────────────────────────────────────────────────────────
    if product_col and sales_col:
        top_prod = df.groupby(product_col)[sales_col].sum().reset_index().sort_values(sales_col, ascending=False).head(10)
        result["top_products"] = top_prod.rename(columns={product_col: "product", sales_col: "sales"}).to_dict("records")

    # ── Top Customers ───────────────────────────────────────────────────────
    if customer_col and sales_col:
        top_cust = df.groupby(customer_col).agg(
            sales=(sales_col, "sum"),
            orders=(sales_col, "count")
        ).reset_index().sort_values("sales", ascending=False).head(10)
        result["top_customers"] = top_cust.rename(columns={customer_col: "customer"}).to_dict("records")

    # ── Regional Performance ────────────────────────────────────────────────
    if region_col and sales_col:
        reg_df = df.groupby(region_col).agg(
            sales=(sales_col, "sum"),
            profit=(profit_col, "sum") if profit_col else (sales_col, "count"),
            orders=(sales_col, "count")
        ).reset_index().sort_values("sales", ascending=False).head(15)
        fig_reg = go.Figure()
        fig_reg.add_trace(go.Bar(
            name="Sales", x=reg_df[region_col], y=reg_df["sales"],
            marker_color="#60a5fa",
            text=reg_df["sales"].apply(lambda v: f"${v:,.0f}"),
            textposition="outside"
        ))
        if profit_col:
            fig_reg.add_trace(go.Bar(
                name="Profit", x=reg_df[region_col], y=reg_df["profit"],
                marker_color="#6ee7b7"
            ))
        fig_reg.update_layout(**_dark_layout("Regional Performance"), barmode="group")
        result["regional_chart"] = fig_to_json(fig_reg)

    # ── Profit vs Loss ──────────────────────────────────────────────────────
    if profit_col:
        pos = float((df[profit_col] > 0).sum())
        neg = float((df[profit_col] <= 0).sum())
        fig_pl = go.Figure(go.Pie(
            labels=["Profitable Orders", "Loss Orders"],
            values=[pos, neg], hole=0.6,
            marker=dict(colors=["#6ee7b7", "#f87171"]),
        ))
        fig_pl.update_layout(**_dark_layout("Profit vs Loss Distribution"))
        result["profit_loss_chart"] = fig_to_json(fig_pl)

    # ── Scatter: Sales vs Profit ────────────────────────────────────────────
    if sales_col and profit_col:
        sample = df[[sales_col, profit_col]].dropna().sample(min(500, len(df)), random_state=42)
        fig_scatter = go.Figure(go.Scatter(
            x=sample[sales_col], y=sample[profit_col],
            mode="markers",
            marker=dict(
                size=6, opacity=0.7,
                color=sample[profit_col],
                colorscale=[[0,"#f87171"],[0.5,"#fbbf24"],[1,"#6ee7b7"]],
                showscale=True,
                colorbar=dict(title="Profit", tickfont=dict(color="#94a3b8"))
            )
        ))
        fig_scatter.update_layout(**_dark_layout("Sales vs Profit Correlation"))
        result["scatter_chart"] = fig_to_json(fig_scatter)

    # ── AI Insights ─────────────────────────────────────────────────────────
    result["insights"] = _generate_insights(df, result["kpis"], sales_col, profit_col, cat_col, region_col, customer_col)
    result["columns"] = list(df.columns)
    result["rows"] = int(df.shape[0])

    return result


def _dark_layout(title: str) -> dict:
    return dict(
        title=dict(text=title, font=dict(color="#e2e8f0", size=14, family="DM Sans")),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#94a3b8", family="DM Sans"),
        margin=dict(l=40, r=20, t=50, b=40),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#94a3b8")),
        xaxis=dict(gridcolor="rgba(148,163,184,0.08)", linecolor="rgba(148,163,184,0.15)", tickfont=dict(color="#64748b")),
        yaxis=dict(gridcolor="rgba(148,163,184,0.08)", linecolor="rgba(148,163,184,0.15)", tickfont=dict(color="#64748b")),
    )


def _generate_insights(df, kpis, sales_col, profit_col, cat_col, region_col, customer_col) -> list:
    insights = []
    if kpis["profit_margin"] > 15:
        insights.append({"type": "positive", "icon": "📈", "text": f"Strong profit margin of {kpis['profit_margin']:.1f}% — well above the 10% e-commerce benchmark."})
    elif kpis["profit_margin"] > 0:
        insights.append({"type": "warning", "icon": "⚠️", "text": f"Profit margin of {kpis['profit_margin']:.1f}% is positive but below 10%. Consider reviewing pricing and COGS."})
    else:
        insights.append({"type": "negative", "icon": "🔻", "text": f"Negative profit margin detected ({kpis['profit_margin']:.1f}%). Urgent review of cost structures recommended."})

    aov = kpis["avg_order_value"]
    insights.append({"type": "info", "icon": "🛒", "text": f"Average order value is ${aov:,.2f}. Upselling or bundling could push this higher."})

    if cat_col and sales_col:
        top_cat = df.groupby(cat_col)[sales_col].sum().idxmax()
        insights.append({"type": "positive", "icon": "🏆", "text": f"'{top_cat}' is your highest-revenue category. Double down on inventory and marketing here."})

    if region_col and sales_col:
        top_region = df.groupby(region_col)[sales_col].sum().idxmax()
        insights.append({"type": "info", "icon": "🌍", "text": f"'{top_region}' is the top-performing region. Explore expansion or deeper penetration there."})

    if profit_col:
        loss_orders = int((df[profit_col] < 0).sum())
        if loss_orders > 0:
            pct = loss_orders / kpis["total_orders"] * 100
            insights.append({"type": "warning", "icon": "💸", "text": f"{loss_orders} orders ({pct:.1f}%) are running at a loss. Investigate discounts and shipping costs on these SKUs."})

    if customer_col and sales_col:
        top5_rev = df.groupby(customer_col)[sales_col].sum().nlargest(5).sum()
        pct = top5_rev / kpis["total_sales"] * 100 if kpis["total_sales"] else 0
        insights.append({"type": "info", "icon": "👑", "text": f"Top 5 customers account for {pct:.1f}% of revenue. Consider a VIP loyalty programme to retain them."})

    return insights


# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    with open("index.html", "r") as f:
        return f.read()


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400

    fname = file.filename.lower()
    try:
        if fname.endswith(".csv"):
            df = pd.read_csv(file, encoding="utf-8", on_bad_lines="skip")
        elif fname.endswith((".xlsx", ".xls")):
            df = pd.read_excel(file)
        else:
            return jsonify({"error": "Unsupported file type. Upload CSV or Excel."}), 400

        df = clean_dataframe(df)
        data = analyse(df)

        # Cache for report generation
        path = os.path.join(UPLOAD_FOLDER, "current.pkl")
        df.to_pickle(path)

        return jsonify({"success": True, "data": data})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/report/excel", methods=["GET"])
def export_excel():
    path = os.path.join(UPLOAD_FOLDER, "current.pkl")
    if not os.path.exists(path):
        return jsonify({"error": "No dataset loaded"}), 400

    df = pd.read_pickle(path)
    df = clean_dataframe(df)
    data = analyse(df)
    kpis = data["kpis"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"

    # Palette
    HDR_FILL = PatternFill("solid", fgColor="0F172A")
    ACCENT = PatternFill("solid", fgColor="10B981")
    ALT = PatternFill("solid", fgColor="1E293B")
    HDR_FONT = Font(bold=True, color="E2E8F0", name="Calibri", size=11)
    TITLE_FONT = Font(bold=True, color="10B981", name="Calibri", size=14)
    BODY_FONT = Font(color="CBD5E1", name="Calibri", size=10)

    def set_row(ws, row, values, fill=None, font=None, bold=False):
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=col, value=val)
            if fill: cell.fill = fill
            if font: cell.font = font
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    # Title
    ws.merge_cells("A1:D1")
    ws["A1"] = "📊  E-Commerce Analytics Report"
    ws["A1"].font = TITLE_FONT
    ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:D2")
    ws["A2"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}   |   Rows: {data['rows']}   |   Columns: {len(data['columns'])}"
    ws["A2"].font = Font(color="64748B", name="Calibri", size=9)
    ws["A2"].fill = HDR_FILL
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 18

    # KPIs
    ws.append([])
    kpi_headers = ["Metric", "Value"]
    set_row(ws, ws.max_row + 1, ["KEY PERFORMANCE INDICATORS", ""], fill=ACCENT, font=Font(bold=True, color="0F172A", size=12))
    set_row(ws, ws.max_row + 1, kpi_headers, fill=HDR_FILL, font=HDR_FONT)
    kpi_rows = [
        ("Total Revenue", f"${kpis['total_sales']:,.2f}"),
        ("Total Profit", f"${kpis['total_profit']:,.2f}"),
        ("Profit Margin", f"{kpis['profit_margin']:.2f}%"),
        ("Total Orders", f"{kpis['total_orders']:,}"),
        ("Total Quantity Sold", f"{kpis['total_quantity']:,}"),
        ("Average Order Value", f"${kpis['avg_order_value']:,.2f}"),
        ("Unique Customers", f"{kpis['unique_customers']:,}"),
    ]
    for i, (k, v) in enumerate(kpi_rows):
        fill = ALT if i % 2 else PatternFill("solid", fgColor="0F172A")
        set_row(ws, ws.max_row + 1, [k, v], fill=fill, font=BODY_FONT)

    # Top Products
    if "top_products" in data:
        ws.append([])
        set_row(ws, ws.max_row + 1, ["TOP PRODUCTS BY REVENUE", ""], fill=ACCENT, font=Font(bold=True, color="0F172A", size=12))
        set_row(ws, ws.max_row + 1, ["Product", "Revenue"], fill=HDR_FILL, font=HDR_FONT)
        for i, row in enumerate(data["top_products"]):
            fill = ALT if i % 2 else PatternFill("solid", fgColor="0F172A")
            set_row(ws, ws.max_row + 1, [row["product"], f"${row['sales']:,.2f}"], fill=fill, font=BODY_FONT)

    # Top Customers
    if "top_customers" in data:
        ws.append([])
        set_row(ws, ws.max_row + 1, ["TOP CUSTOMERS", "", ""], fill=ACCENT, font=Font(bold=True, color="0F172A", size=12))
        set_row(ws, ws.max_row + 1, ["Customer", "Revenue", "Orders"], fill=HDR_FILL, font=HDR_FONT)
        for i, row in enumerate(data["top_customers"]):
            fill = ALT if i % 2 else PatternFill("solid", fgColor="0F172A")
            set_row(ws, ws.max_row + 1, [row["customer"], f"${row['sales']:,.2f}", row["orders"]], fill=fill, font=BODY_FONT)

    # AI Insights
    if "insights" in data:
        ws.append([])
        set_row(ws, ws.max_row + 1, ["AI-GENERATED BUSINESS INSIGHTS", ""], fill=ACCENT, font=Font(bold=True, color="0F172A", size=12))
        for i, ins in enumerate(data["insights"]):
            fill = ALT if i % 2 else PatternFill("solid", fgColor="0F172A")
            set_row(ws, ws.max_row + 1, [ins["icon"] + "  " + ins["text"]], fill=fill, font=BODY_FONT)

    # Adjust column widths
    for col in ws.columns:
        max_w = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_w + 4, 60)

    # Raw data sheet
    ws2 = wb.create_sheet("Raw Data")
    for r in [df.columns.tolist()] + df.head(500).values.tolist():
        ws2.append(r)
    for cell in ws2[1]:
        cell.fill = HDR_FILL
        cell.font = HDR_FONT

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="ecommerce_report.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/sample", methods=["GET"])
def sample_data():
    """Return info about the bundled sample dataset."""
    return jsonify({"message": "Use /upload with a CSV or Excel file. A sample Superstore-style dataset works great."})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
