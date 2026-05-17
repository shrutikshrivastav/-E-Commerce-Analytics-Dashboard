import os
import io
import pandas as pd
import numpy as np
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import warnings
from datetime import datetime
import plotly
import plotly.express as px
warnings.filterwarnings('ignore')

# For PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

app = Flask(__name__)
CORS(app)

# Global variable to store latest analyzed data
latest_analysis = {
    'data': None,
    'raw_df': None,
    'timestamp': None
}

# HTML CONTENT EMBEDDED DIRECTLY
HTML_TEMPLATE = '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
    <title>NexaAnalytics | AI E-Commerce Dashboard</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Inter:opsz,wght@14..32,300;14..32,400;14..32,500;14..32,600;14..32,700&display=swap" rel="stylesheet">
    <script src="https://cdn.plot.ly/plotly-3.0.1.min.js" charset="utf-8"></script>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Inter', sans-serif;
            background: #f0f2f8;
            color: #1e293b;
            overflow-x: hidden;
        }
        .glass-card {
            background: rgba(255, 255, 255, 0.85);
            backdrop-filter: blur(12px);
            border-radius: 32px;
            box-shadow: 0 8px 20px rgba(0, 0, 0, 0.03), 0 2px 4px rgba(0, 0, 0, 0.02);
            border: 1px solid rgba(255, 255, 255, 0.6);
            transition: transform 0.2s ease, box-shadow 0.2s;
        }
        .glass-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 20px 30px -12px rgba(0, 0, 0, 0.1);
            background: rgba(255, 255, 255, 0.92);
        }
        .sidebar {
            background: rgba(18, 25, 45, 0.92);
            backdrop-filter: blur(12px);
            border-right: 1px solid rgba(255,255,255,0.2);
            transition: all 0.3s;
        }
        .sidebar-item {
            display: flex;
            align-items: center;
            gap: 14px;
            padding: 12px 20px;
            margin: 6px 12px;
            border-radius: 24px;
            color: #e2e8f0;
            font-weight: 500;
            transition: 0.2s;
            cursor: pointer;
        }
        .sidebar-item i { width: 24px; font-size: 1.2rem; }
        .sidebar-item.active, .sidebar-item:hover {
            background: rgba(255, 255, 255, 0.15);
            color: white;
            box-shadow: 0 2px 6px rgba(0,0,0,0.1);
        }
        .top-nav {
            background: rgba(255,255,255,0.75);
            backdrop-filter: blur(10px);
            border-bottom: 1px solid rgba(0,0,0,0.05);
        }
        .kpi-value {
            font-size: 2rem;
            font-weight: 700;
            background: linear-gradient(135deg, #1e293b, #2d3a5e);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
        }
        .chart-container { min-height: 380px; width: 100%; }
        .btn-export {
            background: white;
            border: 1px solid #cbd5e1;
            border-radius: 40px;
            padding: 8px 18px;
            font-weight: 500;
            transition: all 0.2s;
            cursor: pointer;
        }
        .btn-export:hover {
            background: #f1f5f9;
            border-color: #94a3b8;
            transform: scale(0.98);
        }
        .ai-insight {
            background: linear-gradient(105deg, #f8fafc 0%, #ffffff 100%);
            border-left: 5px solid #3b82f6;
            border-radius: 24px;
        }
        @media (max-width: 768px) {
            .sidebar {
                position: fixed;
                transform: translateX(-100%);
                z-index: 1050;
                transition: transform 0.2s;
                width: 260px;
            }
            .sidebar.open { transform: translateX(0); }
            .mobile-menu-btn { display: block; }
            .kpi-value { font-size: 1.5rem; }
        }
        .mobile-menu-btn { display: none; }
        @media (max-width: 768px) { .mobile-menu-btn { display: block; } }
    </style>
</head>
<body>
<div class="flex" style="display: flex; min-height: 100vh;">
    <aside class="sidebar" id="sidebar" style="width: 280px; flex-shrink: 0; position: sticky; top:0; height:100vh; overflow-y: auto;">
        <div class="p-6" style="padding: 28px 20px;">
            <div class="flex items-center gap-3 mb-10">
                <i class="fas fa-chart-line text-2xl text-white" style="color:#60a5fa;"></i>
                <span style="font-weight: 700; font-size: 1.5rem; letter-spacing: -0.5px; background: linear-gradient(120deg,#fff,#bfdbfe); -webkit-background-clip:text; background-clip:text; color:transparent;">Nexa<span style="color:#94a3f8;">AI</span></span>
            </div>
            <div class="sidebar-item active" data-section="overview"><i class="fas fa-tachometer-alt"></i><span>Overview</span></div>
            <div class="sidebar-item" data-section="products"><i class="fas fa-box"></i><span>Products & Categories</span></div>
            <div class="sidebar-item" data-section="customers"><i class="fas fa-users"></i><span>Top Customers</span></div>
            <div class="sidebar-item" data-section="regional"><i class="fas fa-map-marker-alt"></i><span>Regional</span></div>
        </div>
    </aside>

    <main style="flex:1; overflow-x: auto;">
        <div class="top-nav px-6 py-4 flex justify-between items-center sticky top-0 z-20" style="padding: 1rem 2rem;">
            <div class="mobile-menu-btn" id="mobileMenuBtn"><i class="fas fa-bars text-2xl text-slate-700"></i></div>
            <h1 class="text-xl font-semibold hidden md:block" style="font-weight:600;">AI-Powered Analytics Studio</h1>
            <div class="flex gap-3">
                <button id="downloadExcelBtn" class="btn-export"><i class="fas fa-file-excel mr-2"></i>Excel Report</button>
                <button id="downloadPdfBtn" class="btn-export"><i class="fas fa-file-pdf mr-2"></i>PDF Report</button>
                <div class="w-8 h-8 rounded-full bg-indigo-100 flex items-center justify-center"><i class="fas fa-user-astronaut text-indigo-700"></i></div>
            </div>
        </div>

        <div class="px-5 md:px-8 py-6" id="dashboardContent">
            <div class="glass-card p-8 mb-8 text-center border-dashed border-2 border-indigo-200">
                <i class="fas fa-cloud-upload-alt fa-3x text-indigo-400 mb-3"></i>
                <h2 class="text-2xl font-semibold">Upload your E‑Commerce Dataset</h2>
                <p class="text-slate-500 mt-1 mb-5">CSV or Excel (orders, sales, profit, customers, categories)</p>
                <div class="flex flex-wrap justify-center gap-4">
                    <label class="bg-indigo-600 hover:bg-indigo-700 text-white px-6 py-2 rounded-full cursor-pointer transition shadow-md">
                        <i class="fas fa-upload mr-2"></i> Choose file
                        <input type="file" id="fileInput" accept=".csv,.xlsx,.xls" style="display:none">
                    </label>
                    <button id="analyzeBtn" class="bg-slate-800 hover:bg-slate-900 text-white px-8 py-2 rounded-full transition shadow-md"><i class="fas fa-chart-simple mr-2"></i>Analyze & Generate</button>
                </div>
                <div id="uploadStatus" class="mt-4 text-sm font-medium"></div>
            </div>

            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5 mb-8" id="kpiRow">
                <div class="glass-card p-5"><div class="text-slate-500 text-sm">Total Revenue</div><div class="kpi-value text-3xl" id="totalRevenue">—</div><div class="text-xs text-green-600 mt-1"><i class="fas fa-trend-up"></i> vs prior period</div></div>
                <div class="glass-card p-5"><div class="text-slate-500 text-sm">Total Profit</div><div class="kpi-value text-3xl" id="totalProfit">—</div><div class="text-xs text-emerald-600">Net margin</div></div>
                <div class="glass-card p-5"><div class="text-slate-500 text-sm">Orders</div><div class="kpi-value text-3xl" id="totalOrders">—</div><div class="text-xs">Units sold</div></div>
                <div class="glass-card p-5"><div class="text-slate-500 text-sm">Avg. Order Value</div><div class="kpi-value text-3xl" id="avgOrderValue">—</div><div class="text-xs">AOV (Revenue/Orders)</div></div>
            </div>

            <div class="glass-card p-5 mb-8 ai-insight">
                <div class="flex items-center gap-2 mb-2"><i class="fas fa-robot text-indigo-500"></i><span class="font-semibold text-lg">AI Business Insights</span></div>
                <div id="aiInsightsText" class="text-slate-700 leading-relaxed">Waiting for dataset upload & analysis. Upload your file and click Analyze.</div>
            </div>

            <div class="grid lg:grid-cols-2 gap-6 mb-8">
                <div class="glass-card p-4"><div class="font-semibold mb-2"><i class="fas fa-chart-line mr-2"></i>Monthly Revenue Trend</div><div id="monthlyChart" class="chart-container"></div></div>
                <div class="glass-card p-4"><div class="font-semibold mb-2"><i class="fas fa-chart-pie mr-2"></i>Sales by Category</div><div id="categoryPieChart" class="chart-container"></div></div>
            </div>

            <div class="grid lg:grid-cols-2 gap-6 mb-8">
                <div class="glass-card p-4"><div class="font-semibold text-lg mb-3"><i class="fas fa-crown text-amber-500 mr-2"></i>Top Selling Products</div><div id="topProductsTable" class="overflow-auto max-h-72"><table class="min-w-full text-sm"><tbody><tr><td class="py-2">—</td></tr></tbody></table></div></div>
                <div class="glass-card p-4"><div class="font-semibold text-lg mb-3"><i class="fas fa-trophy text-amber-500 mr-2"></i>Top Customers (by spend)</div><div id="topCustomersTable" class="overflow-auto max-h-72"><table class="min-w-full text-sm"><tbody><tr><td class="py-2">—</td></tr></tbody></table></div></div>
            </div>

            <div class="grid lg:grid-cols-2 gap-6 mb-8">
                <div class="glass-card p-4"><div class="font-semibold mb-2"><i class="fas fa-globe-americas mr-2"></i>Regional Performance (Sales)</div><div id="regionalChart" class="chart-container"></div></div>
                <div class="glass-card p-4"><div class="font-semibold mb-2"><i class="fas fa-chart-simple mr-2"></i>Profit / Loss by Category</div><div id="profitLossChart" class="chart-container"></div></div>
            </div>
            <div class="text-center text-xs text-slate-400 mt-8 mb-4">Automated analytics powered by Pandas & Plotly — real-time file processing</div>
        </div>
    </main>
</div>

<script>
    let currentDashboardData = null;
    const sidebar = document.getElementById('sidebar');
    document.getElementById('mobileMenuBtn')?.addEventListener('click', () => { sidebar.classList.toggle('open'); });
    document.querySelectorAll('.sidebar-item').forEach(item => {
        item.addEventListener('click', (e) => {
            const section = item.getAttribute('data-section');
            if (section) {
                document.querySelectorAll('.sidebar-item').forEach(i => i.classList.remove('active'));
                item.classList.add('active');
                if(section === 'overview') window.scrollTo({ top: 0, behavior: 'smooth' });
                if(section === 'products') document.getElementById('topProductsTable')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
                if(section === 'customers') document.getElementById('topCustomersTable')?.scrollIntoView({ behavior: 'smooth' });
                if(section === 'regional') document.getElementById('regionalChart')?.scrollIntoView({ behavior: 'smooth' });
            }
            if(window.innerWidth < 768) sidebar.classList.remove('open');
        });
    });

    const fileInput = document.getElementById('fileInput');
    const analyzeBtn = document.getElementById('analyzeBtn');
    const statusDiv = document.getElementById('uploadStatus');

    async function uploadAndAnalyze() {
        if (!fileInput.files.length) {
            statusDiv.innerHTML = '<span class="text-amber-600">⚠️ Please select a CSV or Excel file first.</span>';
            return;
        }
        const file = fileInput.files[0];
        const formData = new FormData();
        formData.append('file', file);
        statusDiv.innerHTML = '<i class="fas fa-spinner fa-pulse"></i> Uploading & analyzing data (AI processing)...';
        try {
            const response = await fetch('/api/analyze', { method: 'POST', body: formData });
            if (!response.ok) throw new Error('Server error');
            const data = await response.json();
            currentDashboardData = data;
            renderDashboard(data);
            statusDiv.innerHTML = '<span class="text-green-600"><i class="fas fa-check-circle"></i> Analysis complete! Dashboard updated.</span>';
        } catch (err) {
            statusDiv.innerHTML = `<span class="text-red-500"><i class="fas fa-exclamation-triangle"></i> Error: ${err.message}</span>`;
        }
    }

    function renderDashboard(data) {
        if (!data) return;
        document.getElementById('totalRevenue').innerText = formatCurrency(data.kpi.total_revenue);
        document.getElementById('totalProfit').innerText = formatCurrency(data.kpi.total_profit);
        document.getElementById('totalOrders').innerText = data.kpi.total_orders?.toLocaleString() || '0';
        document.getElementById('avgOrderValue').innerText = formatCurrency(data.kpi.avg_order_value);
        document.getElementById('aiInsightsText').innerHTML = `<i class="fas fa-lightbulb text-yellow-500 mr-2"></i> ${data.ai_insights}`;

        if (data.monthly_revenue) {
            Plotly.newPlot('monthlyChart', [{
                x: data.monthly_revenue.months, y: data.monthly_revenue.revenues,
                type: 'scatter', mode: 'lines+markers', marker: { color: '#3b82f6', size: 8 },
                line: { width: 3, color: '#2563eb' }, fill: 'tozeroy', fillcolor: 'rgba(59,130,246,0.1)'
            }], { paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)', margin: { t: 30 } });
        }
        if (data.category_sales) {
            Plotly.newPlot('categoryPieChart', [{
                values: Object.values(data.category_sales), labels: Object.keys(data.category_sales),
                type: 'pie', hole: 0.4, textinfo: 'percent+label'
            }], { paper_bgcolor: 'rgba(0,0,0,0)', margin: { t: 20 } });
        }
        if (data.top_products) {
            let html = '<table class="min-w-full"><thead><tr class="border-b"><th>Product</th><th class="text-right">Qty</th><th class="text-right">Revenue</th></tr></thead><tbody>';
            data.top_products.forEach(p => { html += `<tr class="border-b"><td class="py-2">${escapeHtml(p.product)}</td><td class="text-right">${p.quantity}</td><td class="text-right">${formatCurrency(p.revenue)}</td></tr>`; });
            html += '</tbody></table>';
            document.getElementById('topProductsTable').innerHTML = html;
        }
        if (data.top_customers) {
            let html = '<table class="min-w-full"><thead><tr class="border-b"><th>Customer</th><th class="text-right">Spend</th><th class="text-right">Orders</th></tr></thead><tbody>';
            data.top_customers.forEach(c => { html += `<tr class="border-b"><td class="py-2">${escapeHtml(c.customer)}</td><td class="text-right">${formatCurrency(c.spend)}</td><td class="text-right">${c.orders}</td></tr>`; });
            html += '</tbody></table>';
            document.getElementById('topCustomersTable').innerHTML = html;
        }
        if (data.regional_sales) {
            Plotly.newPlot('regionalChart', [{
                x: Object.values(data.regional_sales), y: Object.keys(data.regional_sales),
                type: 'bar', orientation: 'h', marker: { color: '#3b82f6' }
            }], { paper_bgcolor: 'rgba(0,0,0,0)', margin: { l: 100 } });
        }
        if (data.profit_by_category) {
            const cats = Object.keys(data.profit_by_category);
            const profits = Object.values(data.profit_by_category);
            Plotly.newPlot('profitLossChart', [{
                x: cats, y: profits, type: 'bar', marker: { color: profits.map(v => v >= 0 ? '#10b981' : '#ef4444') }
            }], { paper_bgcolor: 'rgba(0,0,0,0)', yaxis: { title: 'Profit ($)' } });
        }
    }

    function formatCurrency(val) { if (!val && val !== 0) return '—'; return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 0 }).format(val); }
    function escapeHtml(str) { if(!str) return ''; return str.replace(/[&<>]/g, function(m){ if(m==='&') return '&amp;'; if(m==='<') return '&lt;'; if(m==='>') return '&gt;'; return m;}); }

    analyzeBtn.addEventListener('click', uploadAndAnalyze);

    async function downloadReport(type) {
        if (!currentDashboardData) { alert('Please upload and analyze a dataset first.'); return; }
        try {
            const response = await fetch(`/api/download_report?type=${type}`);
            if (!response.ok) throw new Error('Report generation failed');
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `ecommerce_analytics.${type === 'excel' ? 'xlsx' : 'pdf'}`;
            document.body.appendChild(a); a.click(); a.remove();
            window.URL.revokeObjectURL(url);
        } catch (err) { alert('Error generating report: ' + err.message); }
    }
    document.getElementById('downloadExcelBtn').addEventListener('click', () => downloadReport('excel'));
    document.getElementById('downloadPdfBtn').addEventListener('click', () => downloadReport('pdf'));
</script>
</body>
</html>'''

@app.route('/')
def index():
    return HTML_TEMPLATE

def format_currency(value):
    return f"${value:,.0f}"

def clean_and_analyze(df):
    """Automatically clean, process and analyze dataset with AI insights"""
    
    # Standardize column names
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
    
    # Column mapping
    revenue_col = None
    profit_col = None
    quantity_col = None
    
    for col in df.columns:
        col_lower = col.lower()
        if 'revenue' in col_lower or 'sales' in col_lower or 'amount' in col_lower:
            revenue_col = col
        elif 'profit' in col_lower:
            profit_col = col
        elif 'quantity' in col_lower or 'qty' in col_lower:
            quantity_col = col
    
    # Convert to numeric
    if revenue_col:
        df[revenue_col] = pd.to_numeric(df[revenue_col], errors='coerce').fillna(0)
    if profit_col:
        df[profit_col] = pd.to_numeric(df[profit_col], errors='coerce').fillna(0)
    elif revenue_col:
        df['estimated_profit'] = df[revenue_col] * 0.3
        profit_col = 'estimated_profit'
    
    if quantity_col:
        df[quantity_col] = pd.to_numeric(df[quantity_col], errors='coerce').fillna(1)
    
    # KPIs
    total_revenue = df[revenue_col].sum() if revenue_col else 0
    total_profit = df[profit_col].sum() if profit_col else 0
    total_orders = len(df)
    avg_order_value = total_revenue / total_orders if total_orders > 0 else 0
    
    # Find date column
    date_col = None
    for col in df.columns:
        if 'date' in col.lower():
            date_col = col
            break
    
    monthly_revenue = {}
    if date_col and revenue_col:
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df['month'] = df[date_col].dt.to_period('M').astype(str)
        monthly_revenue = df.groupby('month')[revenue_col].sum().to_dict()
    
    # Category analysis
    category_col = None
    for col in df.columns:
        if 'category' in col.lower() or 'cat' in col.lower():
            category_col = col
            break
    
    category_sales = {}
    profit_by_category = {}
    if category_col and revenue_col:
        category_sales = df.groupby(category_col)[revenue_col].sum().nlargest(8).to_dict()
        if profit_col:
            profit_by_category = df.groupby(category_col)[profit_col].sum().to_dict()
    
    # Top products
    product_col = None
    for col in df.columns:
        if 'product' in col.lower() or 'item' in col.lower():
            product_col = col
            break
    
    top_products = []
    if product_col and revenue_col:
        qty_col = quantity_col if quantity_col else revenue_col
        product_stats = df.groupby(product_col).agg({revenue_col: 'sum'})
        if quantity_col:
            product_stats[quantity_col] = df.groupby(product_col)[quantity_col].sum()
        product_stats = product_stats.nlargest(5, revenue_col)
        top_products = [
            {'product': idx, 'revenue': row[revenue_col], 'quantity': int(row[quantity_col]) if quantity_col in row else 1}
            for idx, row in product_stats.iterrows()
        ]
    
    # Top customers
    customer_col = None
    for col in df.columns:
        if 'customer' in col.lower() or 'buyer' in col.lower():
            customer_col = col
            break
    
    top_customers = []
    if customer_col and revenue_col:
        customer_stats = df.groupby(customer_col)[revenue_col].sum().nlargest(5)
        customer_orders = df.groupby(customer_col).size()
        top_customers = [
            {'customer': name, 'spend': spend, 'orders': int(customer_orders[name])}
            for name, spend in customer_stats.items()
        ]
    
    # Regional analysis
    region_col = None
    for col in df.columns:
        if 'region' in col.lower() or 'state' in col.lower() or 'country' in col.lower():
            region_col = col
            break
    
    regional_sales = {}
    if region_col and revenue_col:
        regional_sales = df.groupby(region_col)[revenue_col].sum().nlargest(6).to_dict()
    
    # AI Insights generation
    insights = []
    if total_profit > 0:
        profit_margin = (total_profit / total_revenue) * 100 if total_revenue > 0 else 0
        insights.append(f"📊 Profit Margin: {profit_margin:.1f}% — {'Excellent' if profit_margin > 30 else 'Good' if profit_margin > 20 else 'Needs Improvement'} margin business.")
    
    if top_products:
        insights.append(f"🏆 Top Product: '{top_products[0]['product']}' generates {format_currency(top_products[0]['revenue'])} in revenue.")
    
    if top_customers:
        insights.append(f"👥 Top customer '{top_customers[0]['customer']}' contributes {format_currency(top_customers[0]['spend'])} — consider loyalty rewards.")
    
    if category_sales:
        top_cat = max(category_sales, key=category_sales.get)
        insights.append(f"📦 Best-selling category: '{top_cat}' with {format_currency(category_sales[top_cat])} in sales.")
    
    if profit_by_category:
        loss_cats = {k: v for k, v in profit_by_category.items() if v < 0}
        if loss_cats:
            insights.append(f"⚠️ Unprofitable categories: {', '.join(list(loss_cats.keys())[:2])}. Consider pricing optimization.")
        else:
            insights.append(f"✅ All categories are profitable — excellent portfolio management!")
    
    if monthly_revenue and len(monthly_revenue) > 1:
        revenues = list(monthly_revenue.values())
        if revenues[-1] > revenues[0]:
            insights.append(f"📈 Revenue growth trend: +{((revenues[-1]-revenues[0])/revenues[0]*100):.0f}% from first to last month.")
    
    if not insights:
        insights.append("🤖 Upload more detailed data (sales, profit, categories) for deeper AI insights.")
    
    insights.append("💡 Pro tip: Include date columns for trend analysis and region for geographical insights.")
    
    return {
        'kpi': {
            'total_revenue': float(total_revenue),
            'total_profit': float(total_profit),
            'total_orders': int(total_orders),
            'avg_order_value': float(avg_order_value)
        },
        'monthly_revenue': {
            'months': list(monthly_revenue.keys()),
            'revenues': list(monthly_revenue.values())
        } if monthly_revenue else None,
        'category_sales': category_sales,
        'top_products': top_products,
        'top_customers': top_customers,
        'regional_sales': regional_sales,
        'profit_by_category': profit_by_category,
        'ai_insights': ' '.join(insights)
    }

@app.route('/api/analyze', methods=['POST'])
def analyze_file():
    try:
        file = request.files['file']
        if not file:
            return jsonify({'error': 'No file uploaded'}), 400
        
        filename = file.filename.lower()
        if filename.endswith('.csv'):
            df = pd.read_csv(file)
        elif filename.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file)
        else:
            return jsonify({'error': 'Unsupported file type. Use CSV or Excel'}), 400
        
        if df.empty:
            return jsonify({'error': 'Dataset is empty'}), 400
        
        analysis_result = clean_and_analyze(df)
        latest_analysis['data'] = analysis_result
        latest_analysis['raw_df'] = df
        latest_analysis['timestamp'] = datetime.now()
        
        return jsonify(analysis_result)
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/download_report', methods=['GET'])
def download_report():
    report_type = request.args.get('type', 'excel')
    
    if latest_analysis['raw_df'] is None:
        return jsonify({'error': 'No data available. Upload and analyze first.'}), 400
    
    if report_type == 'excel':
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            latest_analysis['raw_df'].to_excel(writer, sheet_name='Raw Data', index=False)
            
            data = latest_analysis['data']
            summary_df = pd.DataFrame([
                ['Total Revenue', data['kpi']['total_revenue']],
                ['Total Profit', data['kpi']['total_profit']],
                ['Total Orders', data['kpi']['total_orders']],
                ['Average Order Value', data['kpi']['avg_order_value']],
                ['AI Insights', data['ai_insights']]
            ], columns=['Metric', 'Value'])
            summary_df.to_excel(writer, sheet_name='KPIs & Insights', index=False)
            
            if data['category_sales']:
                pd.DataFrame(list(data['category_sales'].items()), columns=['Category', 'Sales']).to_excel(writer, sheet_name='Category Sales', index=False)
            if data['top_products']:
                pd.DataFrame(data['top_products']).to_excel(writer, sheet_name='Top Products', index=False)
            if data['top_customers']:
                pd.DataFrame(data['top_customers']).to_excel(writer, sheet_name='Top Customers', index=False)
            if data['regional_sales']:
                pd.DataFrame(list(data['regional_sales'].items()), columns=['Region', 'Sales']).to_excel(writer, sheet_name='Regional Performance', index=False)
        
        output.seek(0)
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f'ecommerce_analytics_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    
    elif report_type == 'pdf':
        output = io.BytesIO()
        doc = SimpleDocTemplate(output, pagesize=letter, title="E-Commerce Analytics Report")
        styles = getSampleStyleSheet()
        story = []
        
        title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=24, textColor=colors.HexColor('#1e293b'), spaceAfter=30)
        story.append(Paragraph("NexaAI E-Commerce Analytics Report", title_style))
        story.append(Spacer(1, 12))
        story.append(Paragraph(f"Generated on {datetime.now().strftime('%B %d, %Y at %H:%M')}", styles['Italic']))
        story.append(Spacer(1, 20))
        
        data = latest_analysis['data']
        kpi_data = [
            ['Metric', 'Value'],
            ['Total Revenue', f"${data['kpi']['total_revenue']:,.0f}"],
            ['Total Profit', f"${data['kpi']['total_profit']:,.0f}"],
            ['Total Orders', f"{data['kpi']['total_orders']:,}"],
            ['Avg Order Value', f"${data['kpi']['avg_order_value']:,.0f}"]
        ]
        kpi_table = Table(kpi_data, colWidths=[200, 150])
        kpi_table.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#3b82f6')), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('GRID', (0,0), (-1,-1), 1, colors.grey)]))
        story.append(kpi_table)
        story.append(Spacer(1, 20))
        
        story.append(Paragraph("🤖 AI Business Insights", styles['Heading2']))
        story.append(Spacer(1, 10))
        story.append(Paragraph(data['ai_insights'], styles['Normal']))
        
        doc.build(story)
        output.seek(0)
        return send_file(output, mimetype='application/pdf', as_attachment=True, download_name=f'ecommerce_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf')
    
    else:
        return jsonify({'error': 'Invalid report type'}), 400

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
