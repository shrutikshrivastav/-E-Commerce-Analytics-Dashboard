from flask import Flask, render_template, request, jsonify
import os

# Initialize Flask
# template_folder='.' tells Flask to look for HTML files in the current directory
app = Flask(__name__, template_folder='.')

@app.route('/')
def dashboard():
    # This will now correctly find index.html in the same folder
    return render_template('index.html')

# ... rest of your code remains the same ...

@app.route('/analyze', methods=['POST'])
def analyze():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    
    try:
        # Read Excel or CSV
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
            
        # --- AI & ANALYTICS LOGIC ---
        
        # 1. Clean Data
        df = df.dropna()
        
        # 2. Basic KPIs (Auto-detect columns)
        sales_col = next((c for c in df.columns if 'sales' in c.lower() or 'revenue' in c.lower()), None)
        profit_col = next((c for c in df.columns if 'profit' in c.lower()), None)
        date_col = next((c for c in df.columns if 'date' in c.lower()), None)
        
        if not sales_col: sales_col = df.columns[0] # Fallback
        
        total_revenue = df[sales_col].sum()
        total_profit = df[profit_col].sum() if profit_col else 0
        
        # 3. Prepare Trend Data (Monthly)
        if date_col:
            df[date_col] = pd.to_datetime(df[date_col])
            df['Month'] = df[date_col].dt.to_period('M')
            monthly_trend = df.groupby('Month')[sales_col].sum().reset_index()
            monthly_trend['Month'] = monthly_trend['Month'].astype(str)
        else:
            monthly_trend = pd.DataFrame(columns=['Month', sales_col])

        # 4. Create Plotly Chart
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=monthly_trend['Month'], 
            y=monthly_trend[sales_col], 
            mode='lines+markers',
            name='Revenue',
            line=dict(color='#6366f1')
        ))
        
        graphJSON = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)
        
        # 5. Simple AI Logic (Rules-based for demo)
        margin = (total_profit / total_revenue) * 100 if total_revenue > 0 else 0
        ai_insight = f"Current profit margin is {margin:.2f}%. " 
        
        if margin > 20:
            ai_insight += "Performance is excellent. Focus on customer retention."
        else:
            ai_insight += "Margins are tight. Review supplier costs."
            
        return jsonify({
            'revenue': float(total_revenue),
            'profit': float(total_profit),
            'chart': graphJSON,
            'insight': ai_insight
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
