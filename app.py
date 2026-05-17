import os
import io
import base64
import pandas as pd
import numpy as np
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
import plotly.express as px
import plotly.graph_objects as go
from plotly.utils import PlotlyJSONEncoder
import json
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# For PDF generation
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.barcharts import VerticalBarChart
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

app = Flask(__name__)
CORS(app)

# Global variable to store latest analyzed data
latest_analysis = {
    'data': None,
    'raw_df': None,
    'timestamp': None
}

def clean_and_analyze(df):
    """Automatically clean, process and analyze dataset with AI insights"""
    
    # Standardize column names
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
    
    # Required column mapping (intelligent detection)
    column_mapping = {
        'revenue': ['revenue', 'sales', 'total_sales', 'amount', 'price', 'sale_amount'],
        'profit': ['profit', 'net_profit', 'gross_profit'],
        'order_id': ['order_id', 'orderid', 'order_number', 'transaction_id'],
        'customer': ['customer', 'customer_name', 'customer_id', 'buyer'],
        'product': ['product', 'product_name', 'item', 'product_id'],
        'category': ['category', 'product_category', 'cat', 'type'],
        'region': ['region', 'state', 'country', 'location', 'area'],
        'quantity': ['quantity', 'qty', 'units', 'order_quantity'],
        'date': ['date', 'order_date', 'transaction_date', 'created_at']
    }
    
    # Map columns
    mapped = {}
    for target, variants in column_mapping.items():
        for col in df.columns:
            if col in variants or any(v in col for v in variants):
                mapped[target] = col
                break
    
    # Basic data cleaning
    if 'date' in mapped:
        df[mapped['date']] = pd.to_datetime(df[mapped['date']], errors='coerce')
        df['month'] = df[mapped['date']].dt.to_period('M').astype(str)
    
    # Calculate revenue if not present
    if 'revenue' not in mapped and 'quantity' in mapped and 'price' in [c for c in df.columns]:
        if 'price' in df.columns:
            df['calculated_revenue'] = df['quantity'] * df['price']
            mapped['revenue'] = 'calculated_revenue'
    
    # Ensure numeric columns
    revenue_col = mapped.get('revenue')
    profit_col = mapped.get('profit')
    quantity_col = mapped.get('quantity')
    
    if revenue_col:
        df[revenue_col] = pd.to_numeric(df[revenue_col], errors='coerce').fillna(0)
    
    if profit_col:
        df[profit_col] = pd.to_numeric(df[profit_col], errors='coerce').fillna(0)
    elif revenue_col:
        # Estimate profit if not available (assuming 30% margin for demo)
        df['estimated_profit'] = df[revenue_col] * 0.3
        profit_col = 'estimated_profit'
        mapped['profit'] = profit_col
    
    if quantity_col:
        df[quantity_col] = pd.to_numeric(df[quantity_col], errors='coerce').fillna(1)
    
    # KPIs
    total_revenue = df[revenue_col].sum() if revenue_col else 0
    total_profit = df[profit_col].sum() if profit_col else 0
    total_orders = len(df) if 'order_id' in mapped else len(df)
    avg_order_value = total_revenue / total_orders if total_orders > 0 else 0
    
    # Monthly Revenue Trend
    monthly_revenue = {}
    if 'month' in df.columns and revenue_col:
        monthly_revenue = df.groupby('month')[revenue_col].sum().to_dict()
    
    # Category-wise Sales
    category_sales = {}
    if 'category' in mapped and revenue_col:
        category_sales = df.groupby(mapped['category'])[revenue_col].sum().nlargest(8).to_dict()
    
    # Top Products
    top_products = []
    if 'product' in mapped and revenue_col and quantity_col:
        product_stats = df.groupby(mapped['product']).agg({
            revenue_col: 'sum',
            quantity_col: 'sum'
        }).rename(columns={revenue_col: 'revenue', quantity_col: 'quantity'})
        product_stats = product_stats.nlargest(5, 'revenue')
        top_products = [
            {'product': idx, 'revenue': row['revenue'], 'quantity': int(row['quantity'])}
            for idx, row in product_stats.iterrows()
        ]
    
    # Top Customers
    top_customers = []
    if 'customer' in mapped and revenue_col:
        customer_stats = df.groupby(mapped['customer'])[revenue_col].sum().nlargest(5)
        customer_orders = df.groupby(mapped['customer']).size()
        top_customers = [
            {'customer': name, 'spend': spend, 'orders': int(customer_orders[name])}
            for name, spend in customer_stats.items()
        ]
    
    # Regional Performance
    regional_sales = {}
    if 'region' in mapped and revenue_col:
        regional_sales = df.groupby(mapped['region'])[revenue_col].sum().nlargest(6).to_dict()
    
    # Profit by Category
    profit_by_category = {}
    if 'category' in mapped and profit_col:
        profit_by_category = df.groupby(mapped['category'])[profit_col].sum().to_dict()
    
    # AI-Generated Business Insights
    insights = []
    if total_profit > 0:
        profit_margin = (total_profit / total_revenue) * 100 if total_revenue > 0 else 0
        insights.append(f"📊 Profit Margin: {profit_margin:.1f}% — {'Healthy' if profit_margin > 25 else 'Moderate' if profit_margin > 15 else 'Low'} margin business.")
    
    if top_products:
        insights.append(f"🏆 Top Product: '{top_products[0]['product']}' generates {format_currency(top_products[0]['revenue'])} in revenue.")
    
    if top_customers:
        insights.append(f"👥 Top customer '{top_customers[0]['customer']}' contributes {format_currency(top_customers[0]['spend'])} — consider loyalty rewards.")
    
    if category_sales and len(category_sales) > 0:
        top_cat = max(category_sales, key=category_sales.get)
        insights.append(f"📦 Best-selling category: '{top_cat}' with {format_currency(category_sales[top_cat])} in sales.")
    
    if profit_by_category:
        loss_cats = {k: v for k, v in profit_by_category.items() if v < 0}
        if loss_cats:
            insights.append(f"⚠️ Unprofitable categories: {', '.join(list(loss_cats.keys())[:2])}. Consider pricing or cost optimization.")
        else:
            insights.append(f"✅ All categories are profitable — excellent portfolio management!")
    
    if monthly_revenue and len(monthly_revenue) > 1:
        revenues = list(monthly_revenue.values())
        if revenues[-1] > revenues[0]:
            insights.append(f"📈 Revenue growth trend: +{((revenues[-1]-revenues[0])/revenues[0]*100):.0f}% from first to last month.")
    
    if not insights:
        insights.append("🤖 Upload more detailed data (sales, profit, categories) for deeper AI insights.")
    
    insights.append("💡 Upload monthly data for trend forecasting and seasonality analysis.")
    
    # Return structured data
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
        'ai_insights': ' '.join(insights),
        'raw_stats': {
            'rows': len(df),
            'columns': len(df.columns),
            'date_range': f"{df['date'].min().date() if 'date' in mapped else 'N/A'} to {df['date'].max().date() if 'date' in mapped else 'N/A'}"
        }
    }

def format_currency(value):
    return f"${value:,.0f}"

@app.route('/')
def index():
    return render_template('index.html')

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
        
        # Process and analyze
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
        return generate_excel_report()
    elif report_type == 'pdf':
        return generate_pdf_report()
    else:
        return jsonify({'error': 'Invalid report type'}), 400

def generate_excel_report():
    """Generate comprehensive Excel report with multiple sheets"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Raw data sheet
        latest_analysis['raw_df'].to_excel(writer, sheet_name='Raw Data', index=False)
        
        # Summary sheet
        summary_data = latest_analysis['data']
        summary_df = pd.DataFrame([
            ['Total Revenue', summary_data['kpi']['total_revenue']],
            ['Total Profit', summary_data['kpi']['total_profit']],
            ['Total Orders', summary_data['kpi']['total_orders']],
            ['Average Order Value', summary_data['kpi']['avg_order_value']],
            ['AI Insights', summary_data['ai_insights']]
        ], columns=['Metric', 'Value'])
        summary_df.to_excel(writer, sheet_name='KPIs & Insights', index=False)
        
        # Category sales
        if summary_data['category_sales']:
            cat_df = pd.DataFrame(list(summary_data['category_sales'].items()), columns=['Category', 'Sales'])
            cat_df.to_excel(writer, sheet_name='Category Sales', index=False)
        
        # Top products
        if summary_data['top_products']:
            prod_df = pd.DataFrame(summary_data['top_products'])
            prod_df.to_excel(writer, sheet_name='Top Products', index=False)
        
        # Top customers
        if summary_data['top_customers']:
            cust_df = pd.DataFrame(summary_data['top_customers'])
            cust_df.to_excel(writer, sheet_name='Top Customers', index=False)
        
        # Regional sales
        if summary_data['regional_sales']:
            reg_df = pd.DataFrame(list(summary_data['regional_sales'].items()), columns=['Region', 'Sales'])
            reg_df.to_excel(writer, sheet_name='Regional Performance', index=False)
    
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'ecommerce_analytics_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    )

def generate_pdf_report():
    """Generate styled PDF report with charts"""
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=letter, title="E-Commerce Analytics Report")
    styles = getSampleStyleSheet()
    story = []
    
    # Title
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=24, textColor=colors.HexColor('#1e293b'), spaceAfter=30)
    story.append(Paragraph("NexaAI E-Commerce Analytics Report", title_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Generated on {datetime.now().strftime('%B %d, %Y at %H:%M')}", styles['Italic']))
    story.append(Spacer(1, 20))
    
    # KPI Table
    data = latest_analysis['data']
    kpi_data = [
        ['Metric', 'Value'],
        ['Total Revenue', f"${data['kpi']['total_revenue']:,.0f}"],
        ['Total Profit', f"${data['kpi']['total_profit']:,.0f}"],
        ['Total Orders', f"{data['kpi']['total_orders']:,}"],
        ['Avg Order Value', f"${data['kpi']['avg_order_value']:,.0f}"]
    ]
    kpi_table = Table(kpi_data, colWidths=[200, 150])
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#3b82f6')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 12),
        ('BOTTOMPADDING', (0,0), (-1,0), 12),
        ('BACKGROUND', (0,1), (-1,-1), colors.beige),
        ('GRID', (0,0), (-1,-1), 1, colors.grey)
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 20))
    
    # AI Insights
    story.append(Paragraph("🤖 AI Business Insights", styles['Heading2']))
    story.append(Spacer(1, 10))
    story.append(Paragraph(data['ai_insights'], styles['Normal']))
    story.append(Spacer(1, 20))
    
    # Top Products
    if data['top_products']:
        story.append(Paragraph("🏆 Top 5 Products", styles['Heading2']))
        prod_data = [['Product', 'Quantity Sold', 'Revenue']]
        for p in data['top_products'][:5]:
            prod_data.append([p['product'], str(p['quantity']), f"${p['revenue']:,.0f}"])
        prod_table = Table(prod_data, colWidths=[200, 80, 100])
        prod_table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.grey), ('ALIGN', (1,0), (-1,-1), 'CENTER')]))
        story.append(prod_table)
    
    doc.build(story)
    output.seek(0)
    return send_file(
        output,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=f'ecommerce_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
