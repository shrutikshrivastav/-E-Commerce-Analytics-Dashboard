from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
import numpy as np
import io
import os
import re
import traceback
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore')

app = Flask(__name__, static_folder='.', template_folder='.')
CORS(app)

CURRENT_DATA = {}

# ══════════════════════════════════════════════════════════════════════════════
#  SMART COLUMN DETECTOR  — understands ANY column name in ANY language/format
# ══════════════════════════════════════════════════════════════════════════════

COLUMN_PATTERNS = {
    'date': {
        'keywords': ['date','dt','time','day','month','year','ordered','purchased',
                     'created','invoice','transaction','sale_date','order_date'],
        'dtype': ['datetime','object'],
        'priority': 1
    },
    'order_id': {
        'keywords': ['order_id','orderid','order id','order no','order number',
                     'invoice','transaction_id','txn','receipt','bill_no','so_no',
                     'reference','ref_no','id'],
        'dtype': ['object','int64','float64'],
        'priority': 2
    },
    'customer': {
        'keywords': ['customer','client','buyer','consumer','account','member',
                     'user','patron','purchaser','cust','name','customer_name',
                     'client_name','ship_to','bill_to'],
        'dtype': ['object'],
        'priority': 3
    },
    'product': {
        'keywords': ['product','item','sku','goods','merchandise','article',
                     'description','prod','commodity','part','model','product_name',
                     'item_name','product_title','listing'],
        'dtype': ['object'],
        'priority': 4
    },
    'category': {
        'keywords': ['category','cat','segment','department','division','class',
                     'type','group','family','genre','catalog','section',
                     'sub_category','subcategory','product_type','line'],
        'dtype': ['object'],
        'priority': 5
    },
    'sub_category': {
        'keywords': ['sub_category','subcategory','sub-category','subcat',
                     'sub_type','sub_segment','sub_class','sub_group'],
        'dtype': ['object'],
        'priority': 6
    },
    'region': {
        'keywords': ['region','state','country','city','location','area',
                     'territory','market','zone','district','province',
                     'geography','geo','place','ship_city','ship_state',
                     'shipping_region','store_location','branch'],
        'dtype': ['object'],
        'priority': 7
    },
    'sales': {
        'keywords': ['sales','revenue','amount','total','turnover','income',
                     'gross_sales','net_sales','sale_amount','order_amount',
                     'total_sales','gmv','value','price','total_price',
                     'subtotal','ext_price','line_total','total_amount',
                     'billing_amount','invoice_amount','receipts'],
        'dtype': ['float64','int64'],
        'priority': 8
    },
    'quantity': {
        'keywords': ['quantity','qty','units','count','pieces','volume',
                     'ordered_qty','sold_qty','units_sold','no_of_units',
                     'pcs','nos','number_of_items','order_qty'],
        'dtype': ['float64','int64'],
        'priority': 9
    },
    'profit': {
        'keywords': ['profit','margin','net_profit','earnings','gain',
                     'net_income','net_margin','profit_amount','contribution',
                     'gross_profit','operating_profit','p_l','pnl'],
        'dtype': ['float64','int64'],
        'priority': 10
    },
    'discount': {
        'keywords': ['discount','disc','rebate','reduction','deduction',
                     'promo','coupon','offer','markdown','allowance'],
        'dtype': ['float64','int64'],
        'priority': 11
    },
    'cost': {
        'keywords': ['cost','cogs','cost_of_goods','purchase_price','unit_cost',
                     'buying_price','landed_cost','expense','expenditure'],
        'dtype': ['float64','int64'],
        'priority': 12
    },
    'ship_mode': {
        'keywords': ['ship_mode','shipping','shipment','delivery','courier',
                     'dispatch','freight','carrier','logistics','transport'],
        'dtype': ['object'],
        'priority': 13
    }
}

def normalize(s):
    """Lowercase, strip, replace spaces/hyphens with underscores."""
    return re.sub(r'[\s\-\.]+', '_', str(s).lower().strip())

def smart_detect_columns(df):
    """
    Intelligent column detection using:
    1. Keyword matching (fuzzy)
    2. Data type validation
    3. Cardinality analysis
    4. Sample value inspection
    """
    col_map = {}
    used_cols = set()
    cols = list(df.columns)
    norm_cols = {normalize(c): c for c in cols}

    def score_column(col, field):
        """Score how well a column matches a field (0-100)."""
        nc = normalize(col)
        score = 0
        pattern = COLUMN_PATTERNS[field]

        # Exact match = 100
        if nc in pattern['keywords']:
            score += 100
        else:
            # Partial match scoring
            for kw in pattern['keywords']:
                if kw in nc:
                    score += 70
                    break
                if nc in kw:
                    score += 50
                    break
                # Character overlap
                common = sum(1 for c in nc if c in kw)
                if common >= min(3, len(nc)-1):
                    overlap = common / max(len(nc), len(kw))
                    if overlap > 0.6:
                        score += int(overlap * 40)

        if score == 0:
            return 0

        # Dtype bonus
        col_dtype = str(df[col].dtype)
        for dt in pattern['dtype']:
            if dt in col_dtype:
                score += 20
                break

        # Cardinality checks
        nunique = df[col].nunique()
        nrows = len(df)

        if field == 'order_id':
            # Should be mostly unique
            if nunique / nrows > 0.8:
                score += 15
        elif field in ('customer', 'product'):
            # Moderate cardinality
            if 2 < nunique < nrows * 0.9:
                score += 10
        elif field in ('category', 'region', 'ship_mode', 'sub_category'):
            # Low cardinality (few distinct values)
            if nunique <= 50:
                score += 15
        elif field in ('sales', 'profit', 'cost'):
            # Numeric, positive mostly
            if pd.api.types.is_numeric_dtype(df[col]):
                score += 20
                if df[col].mean() > 0:
                    score += 5
        elif field == 'quantity':
            if pd.api.types.is_numeric_dtype(df[col]):
                score += 20
                if df[col].min() >= 0 and df[col].max() < 100000:
                    score += 10
        elif field == 'date':
            # Try to parse as date
            sample = df[col].dropna().head(5).astype(str).tolist()
            date_patterns = [r'\d{4}-\d{2}-\d{2}', r'\d{2}/\d{2}/\d{4}',
                           r'\d{2}-\d{2}-\d{4}', r'\d{1,2}/\d{1,2}/\d{2,4}']
            for s in sample:
                for p in date_patterns:
                    if re.search(p, s):
                        score += 25
                        break

        return min(score, 100)

    # Score all columns for all fields, then assign best match
    field_scores = {}
    for field in COLUMN_PATTERNS:
        scores = {}
        for col in cols:
            s = score_column(col, field)
            if s > 0:
                scores[col] = s
        field_scores[field] = scores

    # Greedy assignment: highest confidence first
    sorted_fields = sorted(COLUMN_PATTERNS.keys(),
                          key=lambda f: COLUMN_PATTERNS[f]['priority'])

    for field in sorted_fields:
        scores = field_scores[field]
        available = {c: s for c, s in scores.items() if c not in used_cols}
        if available:
            best_col = max(available, key=available.get)
            if available[best_col] >= 30:  # Minimum confidence threshold
                col_map[field] = best_col
                used_cols.add(best_col)
            else:
                col_map[field] = None
        else:
            col_map[field] = None

    # Fallback: if sales not found, use highest-sum numeric col
    if not col_map.get('sales'):
        num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        num_cols = [c for c in num_cols if c not in used_cols]
        if num_cols:
            best = max(num_cols, key=lambda c: df[c].sum())
            col_map['sales'] = best
            used_cols.add(best)

    # Fallback: profit = second highest numeric if not found
    if not col_map.get('profit') and col_map.get('cost') and col_map.get('sales'):
        # Derive profit = sales - cost
        col_map['profit'] = '__derived_profit__'

    return col_map

def clean_dataframe(df):
    df = df.copy()
    df.columns = df.columns.str.strip()
    df = df.dropna(how='all').drop_duplicates()
    for col in df.select_dtypes(include=[np.number]).columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].fillna('Unknown').astype(str).str.strip()
    return df

def prepare_derived_columns(df, col_map):
    """Create derived columns if needed."""
    if col_map.get('profit') == '__derived_profit__':
        sales_col = col_map['sales']
        cost_col  = col_map['cost']
        df['__profit__'] = df[sales_col] - df[cost_col]
        col_map['profit'] = '__profit__'
    # Parse date
    date_col = col_map.get('date')
    if date_col and date_col in df.columns:
        df[date_col] = pd.to_datetime(df[date_col], infer_datetime_format=True, errors='coerce')
    return df, col_map

def get_col(df, col_map, field):
    """Safely get column values."""
    col = col_map.get(field)
    if col and col in df.columns:
        return df[col]
    return None

# ══════════════════════════════════════════════════════════════════════════════
#  AI INSIGHTS ENGINE
# ══════════════════════════════════════════════════════════════════════════════

def generate_insights(kpis, top_products, category_data, monthly_data,
                      region_data, top_customers):
    insights = []

    rev  = kpis.get('total_revenue', 0)
    prof = kpis.get('total_profit', 0)
    margin = kpis.get('profit_margin', 0)
    aov  = kpis.get('avg_order_value', 0)
    orders = kpis.get('total_orders', 0)

    # 1. Revenue Health
    insights.append({
        "icon": "💰", "type": "revenue",
        "title": "Revenue Health",
        "value": f"${rev:,.0f}",
        "text": f"Total gross revenue is ${rev:,.0f}. " + (
            "Business is performing strongly. Focus on retention to sustain growth."
            if rev > 500000 else
            "Steady revenue base. Expanding product range and marketing can accelerate growth."
        )
    })

    # 2. Profitability
    if margin > 0:
        status = "🟢 Excellent" if margin > 25 else "🟡 Moderate" if margin > 10 else "🔴 Critical"
        insights.append({
            "icon": "📊", "type": "profit",
            "title": "Profitability Analysis",
            "value": f"{margin:.1f}%",
            "text": f"{status} — Profit margin is {margin:.1f}% (${prof:,.0f} net profit). " + (
                "Pricing strategy is working well. Explore premium tiers."
                if margin > 25 else
                "Review cost structure and identify low-margin SKUs for pricing revision."
                if margin > 10 else
                "Urgent: Conduct cost audit. Eliminate or reprice loss-making products immediately."
            )
        })

    # 3. Order Intelligence
    if orders > 0:
        insights.append({
            "icon": "🛒", "type": "orders",
            "title": "Order Intelligence",
            "value": f"{orders:,}",
            "text": f"Processed {orders:,} orders with avg order value of ${aov:,.2f}. " + (
                "High AOV indicates premium customer base. Loyalty programs will retain them."
                if aov > 300 else
                "AOV has room to grow. Bundle offers and upsell prompts can boost cart size."
            )
        })

    # 4. Top Product insight
    if top_products:
        t = top_products[0]
        share = (t['sales'] / rev * 100) if rev > 0 else 0
        insights.append({
            "icon": "🏆", "type": "product",
            "title": "Star Product",
            "value": f"${t['sales']:,.0f}",
            "text": f"'{t['name']}' is the top-grossing product at ${t['sales']:,.0f} ({share:.1f}% of revenue). " +
                    ("High concentration risk — diversify portfolio." if share > 40 else
                     "Healthy contribution. Scale marketing for this SKU.")
        })

    # 5. Category insight
    if len(category_data) >= 2:
        top_cat = category_data[0]
        bot_cat = category_data[-1]
        insights.append({
            "icon": "🗂️", "type": "category",
            "title": "Category Performance",
            "value": top_cat['category'],
            "text": f"'{top_cat['category']}' leads with ${top_cat['sales']:,.0f}. "
                    f"'{bot_cat['category']}' lags at ${bot_cat['sales']:,.0f}. "
                    "Consider cross-promotional strategies between top and bottom categories."
        })

    # 6. Monthly trend
    if len(monthly_data) >= 3:
        revenues = [m['revenue'] for m in monthly_data]
        last3_avg = np.mean(revenues[-3:])
        prev3_avg = np.mean(revenues[-6:-3]) if len(revenues) >= 6 else np.mean(revenues[:3])
        trend_pct = ((last3_avg - prev3_avg) / prev3_avg * 100) if prev3_avg > 0 else 0
        direction = "📈 Upward" if trend_pct > 3 else "📉 Downward" if trend_pct < -3 else "➡️ Flat"
        insights.append({
            "icon": "📅", "type": "trend",
            "title": "Sales Momentum",
            "value": f"{trend_pct:+.1f}%",
            "text": f"{direction} trend over last 3 months vs previous period ({trend_pct:+.1f}%). " + (
                "Momentum is building. Double down on what's working."
                if trend_pct > 3 else
                "Revenue declining. Investigate churn causes and launch win-back campaigns."
                if trend_pct < -3 else
                "Revenue is stable. Identify growth levers to break the plateau."
            )
        })

    # 7. Regional insight
    if len(region_data) >= 2:
        top_reg = region_data[0]
        insights.append({
            "icon": "🌍", "type": "region",
            "title": "Regional Strength",
            "value": top_reg['region'],
            "text": f"'{top_reg['region']}' is the strongest market at ${top_reg['sales']:,.0f}. " +
                    f"Bottom region '{region_data[-1]['region']}' at ${region_data[-1]['sales']:,.0f} needs targeted campaigns."
        })

    # 8. Customer insight
    if top_customers:
        top_cust = top_customers[0]
        cust_share = (top_cust['sales'] / rev * 100) if rev > 0 else 0
        insights.append({
            "icon": "👑", "type": "customer",
            "title": "VIP Customer",
            "value": top_cust['name'],
            "text": f"'{top_cust['name']}' is your highest-value customer at ${top_cust['sales']:,.0f} ({cust_share:.1f}% of revenue). " +
                    ("Create exclusive VIP program to retain top customers." if cust_share > 5 else
                     "Healthy customer distribution. Build loyalty tiers for top 10%.")
        })

    return insights

def generate_executive_summary(kpis, monthly_data, top_products, category_data,
                                 region_data, top_customers, col_map):
    """Generate a full natural language executive summary."""
    rev    = kpis.get('total_revenue', 0)
    prof   = kpis.get('total_profit', 0)
    margin = kpis.get('profit_margin', 0)
    orders = kpis.get('total_orders', 0)
    aov    = kpis.get('avg_order_value', 0)
    custs  = kpis.get('unique_customers', 0)
    prods  = kpis.get('unique_products', 0)
    qty    = kpis.get('total_quantity', 0)

    # Trend analysis
    trend_text = ""
    if len(monthly_data) >= 2:
        revenues = [m['revenue'] for m in monthly_data]
        peak_month = monthly_data[np.argmax(revenues)]['month']
        low_month  = monthly_data[np.argmin(revenues)]['month']
        trend_text = f"Peak revenue month was **{peak_month}** and lowest was **{low_month}**. "

    # Best performers
    top_prod_text = f"Top product: **{top_products[0]['name']}** (${top_products[0]['sales']:,.0f})" if top_products else ""
    top_cat_text  = f"Top category: **{category_data[0]['category']}** (${category_data[0]['sales']:,.0f})" if category_data else ""
    top_reg_text  = f"Top region: **{region_data[0]['region']}** (${region_data[0]['sales']:,.0f})" if region_data else ""
    top_cust_text = f"Top customer: **{top_customers[0]['name']}** (${top_customers[0]['sales']:,.0f})" if top_customers else ""

    # Detected columns info
    detected = [f for f, c in col_map.items() if c and not c.startswith('__')]

    summary = {
        "headline": f"Business generated ${rev:,.0f} in revenue with {margin:.1f}% profit margin across {orders:,} orders",
        "financial": {
            "title": "Financial Performance",
            "points": [
                f"Total Revenue: **${rev:,.0f}**",
                f"Total Profit: **${prof:,.0f}** ({margin:.1f}% margin)",
                f"Average Order Value: **${aov:,.2f}**",
                f"Total Units Sold: **{qty:,}**",
            ]
        },
        "operations": {
            "title": "Operational Metrics",
            "points": [
                f"Total Orders Processed: **{orders:,}**",
                f"Unique Customers: **{custs:,}**",
                f"Unique Products/SKUs: **{prods:,}**",
                top_prod_text, top_cat_text, top_reg_text, top_cust_text
            ]
        },
        "trends": {
            "title": "Trend Analysis",
            "text": trend_text + (
                "Revenue shows positive growth momentum." if len(monthly_data) > 3 and
                monthly_data[-1]['revenue'] > monthly_data[0]['revenue'] else
                "Revenue pattern shows variation — seasonal analysis recommended."
            )
        },
        "recommendations": [
            f"Focus on scaling '{top_products[0]['name'] if top_products else 'top products'}' — highest revenue driver",
            f"'{region_data[-1]['region'] if region_data else 'Underperforming regions'}' needs targeted marketing investment",
            "Implement customer loyalty program for repeat purchase growth",
            f"{'Improve' if margin < 15 else 'Maintain'} profit margins through {'cost reduction' if margin < 15 else 'strategic pricing'}",
            "Expand inventory in top-performing categories to capture more demand",
        ],
        "data_quality": {
            "detected_columns": detected,
            "total_columns_detected": len(detected),
            "column_mapping": {f: c for f, c in col_map.items() if c and not c.startswith('__')}
        }
    }
    return summary

# ══════════════════════════════════════════════════════════════════════════════
#  ROUTES
# ══════════════════════════════════════════════════════════════════════════════

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
        file_bytes = io.BytesIO(file.read())
        filename = file.filename.lower()

        if filename.endswith('.csv'):
            try:
                df = pd.read_csv(file_bytes, encoding='utf-8')
            except:
                file_bytes.seek(0)
                df = pd.read_csv(file_bytes, encoding='latin1')
        elif filename.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_bytes)
        else:
            return jsonify({'error': 'Only CSV and Excel files are supported'}), 400

        df = clean_dataframe(df)
        col_map = smart_detect_columns(df)
        df, col_map = prepare_derived_columns(df, col_map)
        CURRENT_DATA = {'df': df, 'col_map': col_map, 'filename': file.filename}

        # Return detection results so UI can show what was found
        detection_info = {
            f: {'column': c, 'sample': str(df[c].dropna().iloc[0]) if c and c in df.columns else None}
            for f, c in col_map.items() if c
        }

        return jsonify({
            'success': True,
            'rows': len(df),
            'columns': list(df.columns),
            'col_map': col_map,
            'detection_info': detection_info,
            'filename': file.filename
        })
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/analyze', methods=['GET'])
def analyze():
    global CURRENT_DATA
    try:
        if not CURRENT_DATA:
            return jsonify({'error': 'No data loaded'}), 400

        df      = CURRENT_DATA['df']
        col_map = CURRENT_DATA['col_map']

        sc  = col_map.get('sales')
        pc  = col_map.get('profit')
        qc  = col_map.get('quantity')
        prc = col_map.get('product')
        cc  = col_map.get('category')
        rc  = col_map.get('region')
        cuc = col_map.get('customer')
        dc  = col_map.get('date')
        oc  = col_map.get('order_id')
        scc = col_map.get('sub_category')
        dsc = col_map.get('discount')
        shc = col_map.get('ship_mode')

        def s(col): return df[col] if col and col in df.columns else None

        # ── KPIs ──────────────────────────────────────────────────────────────
        total_revenue  = float(df[sc].sum()) if sc else 0
        total_profit   = float(df[pc].sum()) if pc else 0
        total_orders   = int(df[oc].nunique()) if oc else len(df)
        total_qty      = int(df[qc].sum()) if qc else 0
        profit_margin  = (total_profit / total_revenue * 100) if total_revenue > 0 else 0
        avg_order_val  = (total_revenue / total_orders) if total_orders > 0 else 0
        uniq_customers = int(df[cuc].nunique()) if cuc else 0
        uniq_products  = int(df[prc].nunique()) if prc else 0
        total_discount = float(df[dsc].sum()) if dsc else 0
        avg_discount   = float(df[dsc].mean() * 100) if dsc else 0
        return_rate    = 0  # placeholder
        revenue_per_customer = (total_revenue / uniq_customers) if uniq_customers > 0 else 0

        kpis = {
            'total_revenue': total_revenue,
            'total_profit': total_profit,
            'total_orders': total_orders,
            'total_quantity': total_qty,
            'profit_margin': round(profit_margin, 2),
            'avg_order_value': round(avg_order_val, 2),
            'unique_customers': uniq_customers,
            'unique_products': uniq_products,
            'total_discount': round(total_discount, 2),
            'avg_discount_pct': round(avg_discount, 2),
            'revenue_per_customer': round(revenue_per_customer, 2),
        }

        # ── Monthly trend ─────────────────────────────────────────────────────
        monthly_data = []
        if dc and pd.api.types.is_datetime64_any_dtype(df[dc]):
            df['__month__'] = df[dc].dt.to_period('M')
            agg = {'revenue': (sc, 'sum')} if sc else {}
            if pc: agg['profit'] = (pc, 'sum')
            if qc: agg['quantity'] = (qc, 'sum')
            if oc: agg['orders'] = (oc, 'nunique')
            if agg:
                monthly = df.groupby('__month__').agg(**agg).reset_index()
                monthly['__month__'] = monthly['__month__'].astype(str)
                monthly_data = monthly.rename(columns={'__month__': 'month'}).to_dict('records')

        # ── Quarterly trend ───────────────────────────────────────────────────
        quarterly_data = []
        if dc and pd.api.types.is_datetime64_any_dtype(df[dc]) and sc:
            df['__quarter__'] = df[dc].dt.to_period('Q')
            qdf = df.groupby('__quarter__').agg(revenue=(sc,'sum')).reset_index()
            qdf['__quarter__'] = qdf['__quarter__'].astype(str)
            quarterly_data = qdf.rename(columns={'__quarter__':'quarter'}).to_dict('records')

        # ── Top Products ──────────────────────────────────────────────────────
        top_products = []
        if prc and sc:
            agg_dict = {sc: 'sum'}
            if qc: agg_dict[qc] = 'sum'
            if pc: agg_dict[pc] = 'sum'
            tp = df.groupby(prc).agg(agg_dict).reset_index()
            tp = tp.rename(columns={sc:'sales', prc:'name'})
            if qc: tp = tp.rename(columns={qc:'quantity'})
            if pc: tp = tp.rename(columns={pc:'profit'})
            tp = tp.sort_values('sales', ascending=False).head(15)
            top_products = tp.round(2).to_dict('records')

        # ── Top Customers ─────────────────────────────────────────────────────
        top_customers = []
        if cuc and sc:
            agg_dict = {sc: 'sum'}
            if oc: agg_dict[oc] = 'nunique'
            tc = df.groupby(cuc).agg(agg_dict).reset_index()
            tc = tc.rename(columns={sc:'sales', cuc:'name'})
            if oc: tc = tc.rename(columns={oc:'orders'})
            tc = tc.sort_values('sales', ascending=False).head(15)
            top_customers = tc.round(2).to_dict('records')

        # ── Category breakdown ────────────────────────────────────────────────
        category_data = []
        if cc and sc:
            agg_dict = {sc: 'sum'}
            if pc: agg_dict[pc] = 'sum'
            if qc: agg_dict[qc] = 'sum'
            if oc: agg_dict[oc] = 'nunique'
            cat = df.groupby(cc).agg(agg_dict).reset_index()
            cat = cat.rename(columns={sc:'sales', cc:'category'})
            if pc: cat = cat.rename(columns={pc:'profit'})
            if qc: cat = cat.rename(columns={qc:'quantity'})
            if oc: cat = cat.rename(columns={oc:'orders'})
            cat = cat.sort_values('sales', ascending=False)
            category_data = cat.round(2).to_dict('records')

        # ── Sub-category breakdown ────────────────────────────────────────────
        sub_category_data = []
        if scc and sc:
            agg_dict = {sc: 'sum'}
            if pc: agg_dict[pc] = 'sum'
            sc_df = df.groupby(scc).agg(agg_dict).reset_index()
            sc_df = sc_df.rename(columns={sc:'sales', scc:'sub_category'})
            if pc: sc_df = sc_df.rename(columns={pc:'profit'})
            sc_df = sc_df.sort_values('sales', ascending=False).head(15)
            sub_category_data = sc_df.round(2).to_dict('records')

        # ── Regional data ─────────────────────────────────────────────────────
        region_data = []
        if rc and sc:
            agg_dict = {sc: 'sum'}
            if pc: agg_dict[pc] = 'sum'
            if oc: agg_dict[oc] = 'nunique'
            reg = df.groupby(rc).agg(agg_dict).reset_index()
            reg = reg.rename(columns={sc:'sales', rc:'region'})
            if pc: reg = reg.rename(columns={pc:'profit'})
            if oc: reg = reg.rename(columns={oc:'orders'})
            reg = reg.sort_values('sales', ascending=False).head(20)
            region_data = reg.round(2).to_dict('records')

        # ── Ship mode ─────────────────────────────────────────────────────────
        ship_data = []
        if shc and sc:
            sh = df.groupby(shc)[sc].sum().reset_index()
            sh = sh.rename(columns={sc:'sales', shc:'ship_mode'})
            ship_data = sh.sort_values('sales', ascending=False).round(2).to_dict('records')

        # ── Segment / category-region cross ───────────────────────────────────
        cat_region_data = []
        if cc and rc and sc:
            cr = df.groupby([cc, rc])[sc].sum().reset_index()
            cr = cr.rename(columns={sc:'sales', cc:'category', rc:'region'})
            cat_region_data = cr.sort_values('sales', ascending=False).head(30).round(2).to_dict('records')

        # ── AI Insights ───────────────────────────────────────────────────────
        ai_insights = generate_insights(kpis, top_products, category_data,
                                        monthly_data, region_data, top_customers)

        # ── Executive Summary ─────────────────────────────────────────────────
        exec_summary = generate_executive_summary(
            kpis, monthly_data, top_products, category_data,
            region_data, top_customers, col_map
        )

        return jsonify({
            'kpis': kpis,
            'monthly_data': monthly_data,
            'quarterly_data': quarterly_data,
            'top_products': top_products,
            'top_customers': top_customers,
            'category_data': category_data,
            'sub_category_data': sub_category_data,
            'region_data': region_data,
            'ship_data': ship_data,
            'cat_region_data': cat_region_data,
            'ai_insights': ai_insights,
            'exec_summary': exec_summary,
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
        df      = CURRENT_DATA['df']
        col_map = CURRENT_DATA['col_map']
        sc = col_map.get('sales'); pc = col_map.get('profit')
        oc = col_map.get('order_id'); prc = col_map.get('product')
        cuc = col_map.get('customer'); cc = col_map.get('category')

        output = io.BytesIO()
        wb = openpyxl.Workbook()

        hdr  = PatternFill("solid", fgColor="0d1220")
        acc  = PatternFill("solid", fgColor="1a1a3e")
        wf   = Font(color="FFFFFF", bold=True, size=11)
        af   = Font(color="6c63ff", bold=True, size=11)
        center = Alignment(horizontal='center', vertical='center')

        # Sheet 1: Summary
        ws = wb.active; ws.title = "Executive Summary"
        ws.merge_cells('A1:E1')
        ws['A1'] = "E-COMMERCE ANALYTICS — EXECUTIVE REPORT"
        ws['A1'].font = Font(color="6c63ff", bold=True, size=15)
        ws['A1'].alignment = center; ws['A1'].fill = hdr
        ws.row_dimensions[1].height = 32
        ws.merge_cells('A2:E2')
        ws['A2'] = f"Generated: {datetime.now().strftime('%B %d, %Y %H:%M')}   |   File: {CURRENT_DATA.get('filename','')}"
        ws['A2'].font = Font(color="aaaaaa", size=9); ws['A2'].alignment = center; ws['A2'].fill = hdr

        total_revenue = float(df[sc].sum()) if sc else 0
        total_profit  = float(df[pc].sum()) if pc else 0
        total_orders  = int(df[oc].nunique()) if oc else len(df)
        margin = (total_profit/total_revenue*100) if total_revenue else 0

        kpi_list = [
            ("Total Revenue", f"${total_revenue:,.2f}"),
            ("Total Profit",  f"${total_profit:,.2f}"),
            ("Profit Margin", f"{margin:.1f}%"),
            ("Total Orders",  f"{total_orders:,}"),
            ("Avg Order Value", f"${(total_revenue/total_orders if total_orders else 0):,.2f}"),
            ("Unique Customers", f"{int(df[cuc].nunique()) if cuc else 'N/A'}"),
            ("Unique Products",  f"{int(df[prc].nunique()) if prc else 'N/A'}"),
        ]
        ws.append([]); ws.append(["Metric","Value"])
        for cell in ws[ws.max_row]: cell.fill=acc; cell.font=wf; cell.alignment=center
        for label,val in kpi_list:
            ws.append([label, val])
            for cell in ws[ws.max_row]: cell.fill=hdr; cell.font=Font(color="FFFFFF",size=10); cell.alignment=center
        ws.column_dimensions['A'].width = 25; ws.column_dimensions['B'].width = 20

        # Sheet 2: Raw Data
        ws2 = wb.create_sheet("Raw Data")
        for i,col in enumerate(df.columns,1):
            c = ws2.cell(1,i,col); c.fill=hdr; c.font=wf; c.alignment=center
        for _,row in df.head(10000).iterrows():
            ws2.append([str(v) if pd.notna(v) else '' for v in row])
        for i in range(1, len(df.columns)+1):
            ws2.column_dimensions[get_column_letter(i)].width = 16

        # Sheet 3: Top Products
        if prc and sc:
            ws3 = wb.create_sheet("Top Products")
            for i,h in enumerate(["Rank","Product","Revenue","Profit","Qty"],1):
                c=ws3.cell(1,i,h); c.fill=hdr; c.font=wf; c.alignment=center
            tp = df.groupby(prc).agg({sc:'sum',**(
                {pc:'sum'} if pc else {}),**(
                {col_map['quantity']:'sum'} if col_map.get('quantity') else {})
            }).reset_index().sort_values(sc,ascending=False).head(20)
            for rank,(_, row) in enumerate(tp.iterrows(),1):
                ws3.append([rank, str(row[prc]), round(float(row[sc]),2),
                           round(float(row[pc]),2) if pc else 'N/A', 'N/A'])
            for l in ['A','B','C','D','E']: ws3.column_dimensions[l].width=20

        wb.save(output); output.seek(0)
        return send_file(output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True, download_name='analytics_report.xlsx')

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health')
def health():
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
