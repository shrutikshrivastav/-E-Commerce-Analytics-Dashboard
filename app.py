from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.utils import PlotlyJSONEncoder
import json
import io
import os
import traceback
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore')

app = Flask(__name__, static_folder='.', template_folder='.')
CORS(app)

UPLOAD_FOLDER = '/tmp/uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

CURRENT_DATA = {}

def detect_columns(df):
    """Auto-detect column mappings from various CSV formats."""
    col_map = {}
    cols_lower = {c.lower().strip(): c for c in df.columns}

    date_keys = ['date','order date','orderdate','transaction date','sale date','invoice date','purchase date']
    sales_keys = ['sales','revenue','amount','total','order amount','sale amount','total sales','gross sales','net sales','total revenue','price','total price','subtotal']
    profit_keys = ['profit','net profit','profit amount','margin','net income','earnings','net margin']
    quantity_keys = ['quantity','qty','units','count','order qty','quantity ordered','units sold','volume']
    product_keys = ['product','product name','item','item name','product title','sku name','description','product description','goods']
    category_keys = ['category','product category','segment','department','type','product type','sub-category','subcategory','class']
    region_keys = ['region','state','country','city','location','area','territory','market','geography','geo']
    customer_keys = ['customer','customer name','client','buyer','customer id','client name','account','customer_name']
    order_keys = ['order id','orderid','order_id','transaction id','invoice id','invoice no','order no','order number']

    def find_col(keys):
        for k in keys:
            if k in cols_lower:
                return cols_lower[k]
        return None

    col_map['date']     = find_col(date_keys)
    col_map['sales']    = find_col(sales_keys)
    col_map['profit']   = find_col(profit_keys)
    col_map['quantity'] = find_col(quantity_keys)
    col_map['product']  = find_col(product_keys)
    col_map['category'] = find_col(category_keys)
    col_map['region']   = find_col(region_keys)
    col_map['customer'] = find_col(customer_keys)
    col_map['order_id'] = find_col(order_keys)

    # Fallback: use numeric columns for sales/profit
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    if not col_map['sales'] and numeric_cols:
        col_map['sales'] = numeric_cols[0]
    if not col_map['profit'] and len(numeric_cols) > 1:
        col_map['profit'] = numeric_cols[1]

    return col_map

def clean_dataframe(df):
    """Clean and standardize dataframe."""
    df = df.copy()
    df.columns = df.columns.str.strip()
    df = df.dropna(how='all')
    df = df.drop_duplicates()
    for col in df.select_dtypes(include=[np.number]).columns:
        df[col] = df[col].fillna(0)
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].fillna('Unknown').astype(str).str.strip()
    return df

def parse_date_column(df, date_col):
    """Robustly parse date columns."""
    if date_col and date_col in df.columns:
        try:
            df[date_col] = pd.to_datetime(df[date_col], infer_datetime_format=True, errors='coerce')
            df = df.dropna(subset=[date_col])
        except:
            pass
    return df

def generate_ai_insights(kpis, top_products, category_data, monthly_data):
    """Generate rule-based AI business insights."""
    insights = []

    # Revenue insight
    total_revenue = kpis.get('total_revenue', 0)
    if total_revenue > 0:
        insights.append({
            "icon": "💰",
            "title": "Revenue Performance",
            "text": f"Total revenue stands at ${total_revenue:,.0f}. "
                    + ("Strong performance — focus on scaling top channels." if total_revenue > 100000 else "Growth opportunity detected — consider expanding product range.")
        })

    # Profit margin insight
    profit_margin = kpis.get('profit_margin', 0)
    if profit_margin > 0:
        if profit_margin > 25:
            insights.append({"icon": "📈", "title": "Healthy Profit Margin",
                "text": f"Profit margin is {profit_margin:.1f}% — above industry average. Maintain pricing strategy and cost controls."})
        elif profit_margin > 10:
            insights.append({"icon": "⚠️", "title": "Moderate Margin Alert",
                "text": f"Profit margin at {profit_margin:.1f}%. Consider reducing operational costs or revising pricing for low-margin SKUs."})
        else:
            insights.append({"icon": "🔴", "title": "Low Margin Warning",
                "text": f"Profit margin is only {profit_margin:.1f}%. Immediate cost audit recommended — identify and eliminate loss-making products."})

    # Top product insight
    if top_products:
        top = top_products[0]
        insights.append({"icon": "🏆", "title": "Top Performer",
            "text": f"'{top['name']}' leads with ${top['sales']:,.0f} in sales. Double down on this product through upselling and bundle campaigns."})

    # Category insight
    if category_data:
        top_cat = max(category_data, key=lambda x: x['sales'])
        insights.append({"icon": "🗂️", "title": "Category Leader",
            "text": f"'{top_cat['category']}' dominates with ${top_cat['sales']:,.0f} in revenue. Consider expanding inventory in this segment."})

    # Monthly trend
    if len(monthly_data) >= 2:
        last = monthly_data[-1]['revenue']
        prev = monthly_data[-2]['revenue']
        change = ((last - prev) / prev * 100) if prev > 0 else 0
        if change > 5:
            insights.append({"icon": "🚀", "title": "Positive MoM Trend",
                "text": f"Revenue grew {change:.1f}% month-over-month. Momentum is building — invest in marketing to sustain growth."})
        elif change < -5:
            insights.append({"icon": "📉", "title": "Revenue Dip Detected",
                "text": f"Revenue dropped {abs(change):.1f}% vs last month. Investigate demand signals and consider promotional campaigns."})

    # Orders per customer
    avg_order = kpis.get('avg_order_value', 0)
    if avg_order > 0:
        insights.append({"icon": "🛒", "title": "Average Order Value",
            "text": f"AOV is ${avg_order:,.2f}. "
                    + ("Excellent — loyalty programs could push this even higher." if avg_order > 200 else "Consider cross-sell / upsell strategies to increase cart size.")})

    return insights[:6]

@app.route('/')
def index():
    with open('index.html', 'r') as f:
        return f.read()

@app.route('/upload', methods=['POST'])
def upload_file():
    global CURRENT_DATA
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        filename = file.filename.lower()
        file_bytes = io.BytesIO(file.read())

        if filename.endswith('.csv'):
            try:
                df = pd.read_csv(file_bytes, encoding='utf-8')
            except:
                file_bytes.seek(0)
                df = pd.read_csv(file_bytes, encoding='latin1')
        elif filename.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_bytes)
        else:
            return jsonify({'error': 'Only CSV and Excel files supported'}), 400

        df = clean_dataframe(df)
        col_map = detect_columns(df)
        df = parse_date_column(df, col_map['date'])

        CURRENT_DATA = {'df': df, 'col_map': col_map, 'filename': file.filename}

        return jsonify({
            'success': True,
            'rows': len(df),
            'columns': list(df.columns),
            'col_map': col_map,
            'filename': file.filename
        })

    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/analyze', methods=['GET'])
def analyze():
    global CURRENT_DATA
    try:
        if not CURRENT_DATA:
            return jsonify({'error': 'No data loaded. Please upload a file first.'}), 400

        df = CURRENT_DATA['df']
        col_map = CURRENT_DATA['col_map']

        sales_col    = col_map.get('sales')
        profit_col   = col_map.get('profit')
        quantity_col = col_map.get('quantity')
        product_col  = col_map.get('product')
        category_col = col_map.get('category')
        region_col   = col_map.get('region')
        customer_col = col_map.get('customer')
        date_col     = col_map.get('date')
        order_col    = col_map.get('order_id')

        # ── KPIs ──────────────────────────────────────────────────────────────
        total_revenue = float(df[sales_col].sum()) if sales_col else 0
        total_profit  = float(df[profit_col].sum()) if profit_col else 0
        total_orders  = int(df[order_col].nunique()) if order_col else len(df)
        total_qty     = int(df[quantity_col].sum()) if quantity_col else 0
        profit_margin = (total_profit / total_revenue * 100) if total_revenue > 0 else 0
        avg_order_val = (total_revenue / total_orders) if total_orders > 0 else 0
        unique_customers = int(df[customer_col].nunique()) if customer_col else 0
        unique_products  = int(df[product_col].nunique()) if product_col else 0

        kpis = {
            'total_revenue': total_revenue,
            'total_profit': total_profit,
            'total_orders': total_orders,
            'total_quantity': total_qty,
            'profit_margin': round(profit_margin, 2),
            'avg_order_value': round(avg_order_val, 2),
            'unique_customers': unique_customers,
            'unique_products': unique_products
        }

        # ── Monthly Revenue ───────────────────────────────────────────────────
        monthly_data = []
        if date_col and sales_col and pd.api.types.is_datetime64_any_dtype(df[date_col]):
            df['_month'] = df[date_col].dt.to_period('M')
            monthly = df.groupby('_month').agg(
                revenue=(sales_col, 'sum'),
                profit=(profit_col, 'sum') if profit_col else (sales_col, 'count')
            ).reset_index()
            monthly['_month'] = monthly['_month'].astype(str)
            monthly_data = monthly.rename(columns={'_month': 'month'}).to_dict('records')

        # ── Top Products ──────────────────────────────────────────────────────
        top_products = []
        if product_col and sales_col:
            tp = df.groupby(product_col)[sales_col].sum().sort_values(ascending=False).head(10)
            top_products = [{'name': str(k), 'sales': round(float(v), 2)} for k, v in tp.items()]

        # ── Top Customers ─────────────────────────────────────────────────────
        top_customers = []
        if customer_col and sales_col:
            tc = df.groupby(customer_col)[sales_col].sum().sort_values(ascending=False).head(10)
            top_customers = [{'name': str(k), 'sales': round(float(v), 2)} for k, v in tc.items()]

        # ── Category Sales ────────────────────────────────────────────────────
        category_data = []
        if category_col and sales_col:
            cat = df.groupby(category_col).agg(
                sales=(sales_col, 'sum'),
                orders=(order_col, 'nunique') if order_col else (sales_col, 'count')
            ).reset_index()
            cat = cat.sort_values('sales', ascending=False)
            category_data = [{'category': str(r[category_col]), 'sales': round(float(r['sales']), 2),
                               'orders': int(r['orders'])} for _, r in cat.iterrows()]

        # ── Regional Data ─────────────────────────────────────────────────────
        region_data = []
        if region_col and sales_col:
            reg = df.groupby(region_col)[sales_col].sum().sort_values(ascending=False).head(15)
            region_data = [{'region': str(k), 'sales': round(float(v), 2)} for k, v in reg.items()]

        # ── Profit by Category ────────────────────────────────────────────────
        profit_category = []
        if category_col and profit_col:
            pc = df.groupby(category_col)[profit_col].sum().sort_values(ascending=False)
            profit_category = [{'category': str(k), 'profit': round(float(v), 2)} for k, v in pc.items()]

        # ── AI Insights ───────────────────────────────────────────────────────
        ai_insights = generate_ai_insights(kpis, top_products, category_data, monthly_data)

        return jsonify({
            'kpis': kpis,
            'monthly_data': monthly_data,
            'top_products': top_products,
            'top_customers': top_customers,
            'category_data': category_data,
            'region_data': region_data,
            'profit_category': profit_category,
            'ai_insights': ai_insights,
            'col_map': col_map,
            'filename': CURRENT_DATA.get('filename', '')
        })

    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/download/excel', methods=['GET'])
def download_excel():
    global CURRENT_DATA
    try:
        if not CURRENT_DATA:
            return jsonify({'error': 'No data loaded'}), 400

        df = CURRENT_DATA['df']
        col_map = CURRENT_DATA['col_map']

        output = io.BytesIO()
        wb = openpyxl.Workbook()

        # ── Summary Sheet ─────────────────────────────────────────────────────
        ws = wb.active
        ws.title = "Executive Summary"

        header_fill  = PatternFill("solid", fgColor="1a1a2e")
        accent_fill  = PatternFill("solid", fgColor="0f3460")
        kpi_fill     = PatternFill("solid", fgColor="16213e")
        white_font   = Font(color="FFFFFF", bold=True, size=12)
        accent_font  = Font(color="e94560", bold=True, size=11)
        normal_font  = Font(color="FFFFFF", size=10)
        center_align = Alignment(horizontal='center', vertical='center')

        ws.merge_cells('A1:F1')
        ws['A1'] = "🛒 E-Commerce Analytics — Executive Report"
        ws['A1'].font = Font(color="e94560", bold=True, size=16)
        ws['A1'].alignment = center_align
        ws['A1'].fill = header_fill
        ws.row_dimensions[1].height = 35

        ws.merge_cells('A2:F2')
        ws['A2'] = f"Generated: {datetime.now().strftime('%B %d, %Y %H:%M')} | Source: {CURRENT_DATA.get('filename','')}"
        ws['A2'].font = Font(color="aaaaaa", size=10)
        ws['A2'].alignment = center_align
        ws['A2'].fill = header_fill

        # KPI section
        sales_col   = col_map.get('sales')
        profit_col  = col_map.get('profit')
        order_col   = col_map.get('order_id')
        product_col = col_map.get('product')
        customer_col= col_map.get('customer')

        total_revenue = float(df[sales_col].sum()) if sales_col else 0
        total_profit  = float(df[profit_col].sum()) if profit_col else 0
        total_orders  = int(df[order_col].nunique()) if order_col else len(df)
        margin        = (total_profit/total_revenue*100) if total_revenue else 0

        kpi_rows = [
            ("Total Revenue",   f"${total_revenue:,.2f}"),
            ("Total Profit",    f"${total_profit:,.2f}"),
            ("Profit Margin",   f"{margin:.1f}%"),
            ("Total Orders",    f"{total_orders:,}"),
            ("Unique Customers",f"{int(df[customer_col].nunique()) if customer_col else 'N/A'}"),
            ("Unique Products",  f"{int(df[product_col].nunique()) if product_col else 'N/A'}"),
        ]

        ws.append([])
        ws.append(["KPI", "Value"])
        for cell in ws[ws.max_row]:
            cell.fill = accent_fill
            cell.font = white_font
            cell.alignment = center_align

        for label, value in kpi_rows:
            ws.append([label, value])
            for cell in ws[ws.max_row]:
                cell.fill = kpi_fill
                cell.font = normal_font
                cell.alignment = center_align

        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 20

        # ── Raw Data Sheet ────────────────────────────────────────────────────
        ws2 = wb.create_sheet("Raw Data")
        headers = list(df.columns)
        ws2.append(headers)
        for cell in ws2[1]:
            cell.fill = header_fill
            cell.font = white_font
            cell.alignment = center_align

        for _, row in df.head(5000).iterrows():
            ws2.append([str(v) if pd.notna(v) else '' for v in row])

        for i, col in enumerate(headers, 1):
            ws2.column_dimensions[get_column_letter(i)].width = 18

        # ── Top Products Sheet ────────────────────────────────────────────────
        if col_map.get('product') and col_map.get('sales'):
            ws3 = wb.create_sheet("Top Products")
            ws3.append(["Rank", "Product", "Revenue", "% Share"])
            for cell in ws3[1]:
                cell.fill = header_fill; cell.font = white_font; cell.alignment = center_align
            tp = df.groupby(col_map['product'])[col_map['sales']].sum().sort_values(ascending=False).head(20)
            for i, (k, v) in enumerate(tp.items(), 1):
                pct = v / total_revenue * 100 if total_revenue else 0
                ws3.append([i, str(k), round(float(v), 2), f"{pct:.1f}%"])
            for col_letter in ['A','B','C','D']:
                ws3.column_dimensions[col_letter].width = 20

        wb.save(output)
        output.seek(0)
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name='ecommerce_analytics_report.xlsx')

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'version': '1.0.0'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
