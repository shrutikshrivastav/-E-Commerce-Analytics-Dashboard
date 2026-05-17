from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import numpy as np
import io
from datetime import datetime
import plotly.express as px
import plotly.io as pio

app = Flask(__name__)

# Store uploaded data in memory
current_df = None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    global current_df
    file = request.files['file']
    if not file:
        return jsonify({"error": "No file uploaded"}), 400

    # Read CSV or Excel dynamically
    if file.filename.endswith('.csv'):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    # Clean & enrich
    df.columns = [c.strip().lower().replace(' ', '_') for c in df.columns]
    df.dropna(how='all', inplace=True)

    # Ensure required columns exist
    required = ['order_date', 'customer', 'product', 'category', 'sales', 'profit']
    for col in required:
        if col not in df.columns:
            return jsonify({"error": f"Missing required column: {col}"}), 400

    # Convert types
    df['order_date'] = pd.to_datetime(df['order_date'], errors='coerce')
    df['month'] = df['order_date'].dt.to_period('M').astype(str)
    df['region'] = df.get('region', 'Unknown')

    current_df = df

    return jsonify({"message": "File uploaded and processed successfully!"})

@app.route('/data', methods=['GET'])
def get_data():
    global current_df
    if current_df is None:
        return jsonify({"error": "No data loaded yet"}), 400

    df = current_df
    total_sales = round(df['sales'].sum(), 2)
    total_profit = round(df['profit'].sum(), 2)
    top_products = df.groupby('product')['sales'].sum().nlargest(5).to_dict()
    top_customers = df.groupby('customer')['sales'].sum().nlargest(5).to_dict()
    monthly_revenue = df.groupby('month')['sales'].sum().reset_index()

    trend_chart = pio.to_json(px.line(monthly_revenue, x='month', y='sales', title='Monthly Revenue Trend'))
    category_chart = pio.to_json(px.bar(df.groupby('category')['sales'].sum().reset_index(),
                                        x='category', y='sales',
                                        title='Sales by Category'))

    insights = generate_ai_insights(df)

    return jsonify({
        "total_sales": total_sales,
        "total_profit": total_profit,
        "top_products": top_products,
        "top_customers": top_customers,
        "trend_chart": trend_chart,
        "category_chart": category_chart,
        "insights": insights
    })

def generate_ai_insights(df):
    growth = round(df['sales'].pct_change().mean() * 100, 2)
    best_cat = df.groupby('category')['profit'].sum().idxmax()
    worst_cat = df.groupby('category')['profit'].sum().idxmin()

    return f"Sales grew approximately {growth}% month-over-month. " \
           f"The most profitable category is '{best_cat}' and the least profitable is '{worst_cat}'."

@app.route('/download', methods=['GET'])
def download_excel():
    global current_df
    if current_df is None:
        return jsonify({"error": "No data loaded yet"}), 400

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        current_df.to_excel(writer, index=False, sheet_name='E-Commerce Data')

        summary = pd.DataFrame({
            'Metric': ['Total Sales', 'Total Profit', 'Record Count'],
            'Value': [current_df['sales'].sum(), current_df['profit'].sum(), len(current_df)]
        })
        summary.to_excel(writer, index=False, sheet_name='Summary')

    output.seek(0)
    return send_file(output, as_attachment=True, download_name='Ecommerce_Report.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
