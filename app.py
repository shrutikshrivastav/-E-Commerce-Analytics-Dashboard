import os
import io
import json
import traceback
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__, static_folder=".", template_folder=".")
CORS(app)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

COLORS = ["#6ee7b7","#f472b6","#60a5fa","#fbbf24","#a78bfa","#fb923c","#38bdf8","#4ade80","#e879f9","#f97316"]

def fig_to_json(fig):
    return json.loads(fig.to_json())

def safe_col(df, *candidates):
    lower = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c.lower() in lower:
            return lower[c.lower()]
    return None

def dark_layout(title="", height=300):
    return dict(
        title=dict(text=title, font=dict(color="#e2e8f0", size=13, family="DM Sans"), x=0.01),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#94a3b8", family="DM Sans", size=11),
        height=height,
        margin=dict(l=40, r=20, t=44, b=36),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#94a3b8", size=10)),
        xaxis=dict(gridcolor="rgba(148,163,184,0.07)", linecolor="rgba(148,163,184,0.12)",
                   tickfont=dict(color="#64748b", size=10), zeroline=False),
        yaxis=dict(gridcolor="rgba(148,163,184,0.07)", linecolor="rgba(148,163,184,0.12)",
                   tickfont=dict(color="#64748b", size=10), zeroline=False),
    )

def clean_dataframe(df):
    df.columns = df.columns.str.strip()
    df = df.dropna(how="all").reset_index(drop=True)
    mapping = {
        "order_id":["orderid","order id","id","transaction_id","order_number"],
        "date":["order_date","orderdate","sale_date","purchase_date","invoice_date","ship_date"],
        "sales":["revenue","amount","total","sale_amount","total_sales","price","gmv"],
        "profit":["profit_amount","net_profit","margin","earnings"],
        "quantity":["qty","units","quantity_ordered","unit_sales"],
        "category":["product_category","cat","segment","dept","department"],
        "sub_category":["sub_cat","subcategory","sub category","sub-category","product_type"],
        "product_name":["product","item","item_name","product name","sku_name","product_title"],
        "customer_name":["customer","client","buyer","customer name","cust_name","customer_id"],
        "region":["area","zone","territory","state","country","location","geo"],
        "city":["town","city_name"],
        "discount":["disc","discount_pct","discount_rate"],
        "ship_mode":["shipping_mode","delivery_mode","ship mode","shipping"],
    }
    lower_cols = {c.lower(): c for c in df.columns}
    renames = {}
    for canonical, synonyms in mapping.items():
        if canonical not in lower_cols:
            for syn in synonyms:
                if syn in lower_cols:
                    renames[lower_cols[syn]] = canonical
                    break
    if renames:
        df = df.rename(columns=renames)
    date_col = safe_col(df, "date", "order_date")
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], infer_datetime_format=True, errors="coerce")
    for col in ["sales","profit","quantity","discount"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df

def load_df():
    path = os.path.join(UPLOAD_FOLDER, "current.pkl")
    if not os.path.exists(path):
        return None
    return pd.read_pickle(path)

def _insights(df, kpis, sales_col, profit_col, cat_col, region_col, customer_col, qty_col):
    ins = []
    pm = kpis.get("profit_margin", 0)
    if pm > 15:
        ins.append({"type":"positive","icon":"📈","text":f"Strong profit margin of {pm:.1f}% — well above the 10% e-commerce benchmark. Excellent cost management."})
    elif pm > 0:
        ins.append({"type":"warning","icon":"⚠️","text":f"Profit margin is {pm:.1f}% — positive but below 10% benchmark. Review pricing strategy and COGS."})
    else:
        ins.append({"type":"negative","icon":"🔻","text":f"Negative profit margin of {pm:.1f}%. Immediate review of cost structure and pricing is critical."})
    aov = kpis.get("avg_order_value", 0)
    ins.append({"type":"info","icon":"🛒","text":f"Average order value is ${aov:,.2f}. Product bundling, upselling, and free-shipping thresholds can increase this KPI."})
    if cat_col and sales_col:
        top_cat = df.groupby(cat_col)[sales_col].sum().idxmax()
        top_val = df.groupby(cat_col)[sales_col].sum().max()
        share = top_val / kpis["total_sales"] * 100 if kpis["total_sales"] else 0
        ins.append({"type":"positive","icon":"🏆","text":f"'{top_cat}' leads with {share:.1f}% of total revenue (${top_val:,.0f}). Strengthen inventory and marketing here."})
    if region_col and sales_col:
        top_reg = df.groupby(region_col)[sales_col].sum().idxmax()
        ins.append({"type":"info","icon":"🌍","text":f"'{top_reg}' is the highest-revenue region. Evaluate market saturation before further spend there."})
    if profit_col:
        loss = kpis.get("loss_orders", 0)
        if loss > 0:
            pct = loss / kpis["total_orders"] * 100
            ins.append({"type":"warning","icon":"💸","text":f"{loss:,} orders ({pct:.1f}%) are unprofitable. Audit discounts, returns, and shipping costs on these transactions."})
    if customer_col and sales_col:
        top5 = df.groupby(customer_col)[sales_col].sum().nlargest(5).sum()
        pct5 = top5 / kpis["total_sales"] * 100 if kpis["total_sales"] else 0
        ins.append({"type":"info","icon":"👑","text":f"Top 5 customers represent {pct5:.1f}% of total revenue — high concentration risk. A VIP retention programme is recommended."})
    total_cust = kpis.get("unique_customers", 0)
    total_orders = kpis.get("total_orders", 0)
    if total_cust > 0:
        freq = total_orders / total_cust
        ins.append({"type":"info","icon":"🔄","text":f"Average purchase frequency is {freq:.1f} orders per customer. Loyalty programmes can push this above 3x."})
    return ins

def analyse_overview(df):
    date_col=safe_col(df,"date","order_date"); sales_col=safe_col(df,"sales","revenue","amount")
    profit_col=safe_col(df,"profit"); qty_col=safe_col(df,"quantity","qty")
    cat_col=safe_col(df,"category","product_category"); customer_col=safe_col(df,"customer_name","customer")
    region_col=safe_col(df,"region","state","country")
    total_sales=float(df[sales_col].sum()) if sales_col else 0
    total_profit=float(df[profit_col].sum()) if profit_col else 0
    total_orders=int(df.shape[0]); total_qty=int(df[qty_col].sum()) if qty_col else 0
    aov=total_sales/total_orders if total_orders else 0
    margin=(total_profit/total_sales*100) if total_sales else 0
    unique_cust=int(df[customer_col].nunique()) if customer_col else 0
    loss_orders=int((df[profit_col]<0).sum()) if profit_col else 0
    kpis=dict(total_sales=round(total_sales,2),total_profit=round(total_profit,2),total_orders=total_orders,
              total_quantity=total_qty,avg_order_value=round(aov,2),profit_margin=round(margin,2),
              unique_customers=unique_cust,loss_orders=loss_orders)
    charts={}
    if date_col and sales_col:
        grp=df.dropna(subset=[date_col]).copy(); grp["_m"]=grp[date_col].dt.to_period("M")
        agg_d={"sales":(sales_col,"sum")}
        if profit_col: agg_d["profit"]=(profit_col,"sum")
        m=grp.groupby("_m").agg(**agg_d).reset_index(); m["_m"]=m["_m"].astype(str)
        fig=go.Figure()
        fig.add_trace(go.Scatter(x=m["_m"],y=m["sales"],name="Revenue",mode="lines+markers",
            line=dict(color="#6ee7b7",width=2.5),marker=dict(size=5),fill="tozeroy",fillcolor="rgba(110,231,183,0.07)"))
        if profit_col:
            fig.add_trace(go.Scatter(x=m["_m"],y=m["profit"],name="Profit",mode="lines+markers",
                line=dict(color="#f472b6",width=2,dash="dot"),marker=dict(size=4)))
        fig.update_layout(**dark_layout("Monthly Revenue & Profit",300)); charts["monthly_trend"]=fig_to_json(fig)
    if cat_col and sales_col:
        cd=df.groupby(cat_col)[sales_col].sum().reset_index().sort_values(sales_col,ascending=False)
        fig=go.Figure(go.Pie(labels=cd[cat_col],values=cd[sales_col],hole=0.58,
            marker=dict(colors=COLORS),textfont=dict(size=11),
            hovertemplate="<b>%{label}</b><br>$%{value:,.0f}<br>%{percent}<extra></extra>"))
        fig.update_layout(**dark_layout("Sales by Category",300)); charts["category_donut"]=fig_to_json(fig)
    if region_col and sales_col:
        rd=df.groupby(region_col)[sales_col].sum().reset_index().sort_values(sales_col,ascending=False).head(8)
        fig=go.Figure(go.Bar(x=rd[region_col],y=rd[sales_col],marker=dict(color=COLORS[:len(rd)]),
            text=rd[sales_col].apply(lambda v:f"${v/1e3:.1f}K"),textposition="outside"))
        fig.update_layout(**dark_layout("Revenue by Region",260)); charts["region_bar"]=fig_to_json(fig)
    if profit_col:
        pos=float((df[profit_col]>0).sum()); neg=float((df[profit_col]<=0).sum())
        fig=go.Figure(go.Pie(labels=["Profitable","Loss"],values=[pos,neg],hole=0.6,
            marker=dict(colors=["#6ee7b7","#f87171"])))
        fig.update_layout(**dark_layout("Profit vs Loss Orders",280)); charts["profit_loss"]=fig_to_json(fig)
    insights=_insights(df,kpis,sales_col,profit_col,cat_col,region_col,customer_col,qty_col)
    return dict(kpis=kpis,charts=charts,insights=insights,columns=list(df.columns),rows=int(df.shape[0]))

def analyse_trends(df):
    date_col=safe_col(df,"date","order_date"); sales_col=safe_col(df,"sales","revenue","amount")
    profit_col=safe_col(df,"profit"); qty_col=safe_col(df,"quantity","qty")
    charts={}
    if not date_col or not sales_col:
        return {"charts":charts,"error":"No date or sales column found"}
    grp=df.dropna(subset=[date_col]).copy()
    grp["_month"]=grp[date_col].dt.to_period("M"); grp["_quarter"]=grp[date_col].dt.to_period("Q")
    grp["_year"]=grp[date_col].dt.year; grp["_dow"]=grp[date_col].dt.day_name()
    agg={"sales":(sales_col,"sum")}
    if profit_col: agg["profit"]=(profit_col,"sum")
    if qty_col: agg["qty"]=(qty_col,"sum")
    m=grp.groupby("_month").agg(**agg).reset_index(); m["_month"]=m["_month"].astype(str)
    fig=go.Figure()
    fig.add_trace(go.Bar(x=m["_month"],y=m["sales"],name="Revenue",marker_color="rgba(110,231,183,0.7)",
        hovertemplate="<b>%{x}</b><br>$%{y:,.0f}<extra></extra>"))
    if profit_col:
        ly2=dict(overlaying="y",side="right",showgrid=False,tickfont=dict(color="#f472b6",size=10),zeroline=False)
        fig.add_trace(go.Scatter(x=m["_month"],y=m["profit"],name="Profit",mode="lines+markers",
            line=dict(color="#f472b6",width=2),yaxis="y2"))
    lyt=dark_layout("Monthly Revenue vs Profit",320)
    if profit_col: lyt["yaxis2"]=ly2
    fig.update_layout(**lyt); charts["monthly"]=fig_to_json(fig)
    q=grp.groupby("_quarter").agg(**agg).reset_index(); q["_quarter"]=q["_quarter"].astype(str)
    fig2=go.Figure()
    fig2.add_trace(go.Bar(x=q["_quarter"],y=q["sales"],name="Revenue",
        marker=dict(color=COLORS[:len(q)]),text=q["sales"].apply(lambda v:f"${v/1e3:.1f}K"),textposition="outside"))
    if profit_col:
        fig2.add_trace(go.Scatter(x=q["_quarter"],y=q["profit"],name="Profit",
            mode="lines+markers",line=dict(color="#f472b6",width=2)))
    fig2.update_layout(**dark_layout("Quarterly Performance",300)); charts["quarterly"]=fig_to_json(fig2)
    yr=grp.groupby("_year").agg(**agg).reset_index()
    fig3=go.Figure()
    fig3.add_trace(go.Bar(x=yr["_year"].astype(str),y=yr["sales"],name="Revenue",marker_color="#60a5fa",
        text=yr["sales"].apply(lambda v:f"${v/1e3:.1f}K"),textposition="outside"))
    if profit_col:
        fig3.add_trace(go.Bar(x=yr["_year"].astype(str),y=yr["profit"],name="Profit",marker_color="#6ee7b7"))
    fig3.update_layout(**dark_layout("Year-over-Year Revenue",280),barmode="group"); charts["yearly"]=fig_to_json(fig3)
    dow_order=["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    dows=grp.groupby("_dow")[sales_col].mean().reindex(dow_order).fillna(0)
    fig4=go.Figure(go.Bar(x=dows.index,y=dows.values,
        marker=dict(color=dows.values,colorscale=[[0,"#1e3a5f"],[0.5,"#3b82f6"],[1,"#6ee7b7"]],showscale=False),
        text=[f"${v:,.0f}" for v in dows.values],textposition="outside"))
    fig4.update_layout(**dark_layout("Avg Revenue by Day of Week",260)); charts["dow"]=fig_to_json(fig4)
    if len(m)>=2:
        m["growth"]=m["sales"].pct_change()*100
        cg=["#f87171" if v<0 else "#6ee7b7" for v in m["growth"].fillna(0)]
        fig5=go.Figure(go.Bar(x=m["_month"],y=m["growth"].fillna(0),marker_color=cg,
            text=m["growth"].fillna(0).apply(lambda v:f"{v:+.1f}%"),textposition="outside"))
        fig5.update_layout(**dark_layout("Month-over-Month Growth %",260)); charts["mom_growth"]=fig_to_json(fig5)
    summary={}
    if len(m)>=1:
        summary["best_month"]=m.loc[m["sales"].idxmax(),"_month"]
        summary["best_revenue"]=float(m["sales"].max())
        summary["avg_monthly"]=float(m["sales"].mean())
    return {"charts":charts,"summary":summary}

def analyse_categories(df):
    sales_col=safe_col(df,"sales","revenue","amount"); profit_col=safe_col(df,"profit")
    cat_col=safe_col(df,"category","product_category"); subcat_col=safe_col(df,"sub_category","subcategory")
    charts={}; table_data=[]
    if cat_col and sales_col:
        agg_d={"sales":(sales_col,"sum"),"orders":(sales_col,"count")}
        if profit_col: agg_d["profit"]=(profit_col,"sum")
        cd=df.groupby(cat_col).agg(**agg_d).reset_index().sort_values("sales",ascending=False)
        fig=go.Figure(go.Pie(labels=cd[cat_col],values=cd["sales"],hole=0.55,
            marker=dict(colors=COLORS),textinfo="label+percent",
            hovertemplate="<b>%{label}</b><br>$%{value:,.0f}<br>%{percent}<extra></extra>"))
        fig.update_layout(**dark_layout("Revenue Share by Category",310)); charts["cat_donut"]=fig_to_json(fig)
        fig2=go.Figure(go.Bar(x=cd["sales"],y=cd[cat_col],orientation="h",
            marker=dict(color=COLORS[:len(cd)]),text=cd["sales"].apply(lambda v:f"${v:,.0f}"),textposition="outside"))
        fig2.update_layout(**dark_layout("Revenue by Category",280)); charts["cat_bar"]=fig_to_json(fig2)
        if profit_col and "profit" in cd.columns:
            cd["margin"]=(cd["profit"]/cd["sales"]*100).round(1)
            fig3=go.Figure(go.Bar(x=cd[cat_col],y=cd["margin"],
                marker_color=["#6ee7b7" if v>0 else "#f87171" for v in cd["margin"]],
                text=cd["margin"].apply(lambda v:f"{v:.1f}%"),textposition="outside"))
            fig3.update_layout(**dark_layout("Profit Margin % by Category",280)); charts["cat_margin"]=fig_to_json(fig3)
        table_data=cd.rename(columns={cat_col:"category"}).to_dict("records")
    if subcat_col and sales_col:
        agg_d2={"sales":(sales_col,"sum"),"orders":(sales_col,"count")}
        if profit_col: agg_d2["profit"]=(profit_col,"sum")
        sd=df.groupby(subcat_col).agg(**agg_d2).reset_index().sort_values("sales",ascending=False).head(15)
        fig4=go.Figure(go.Bar(x=sd["sales"],y=sd[subcat_col],orientation="h",
            marker=dict(color=sd["sales"],colorscale=[[0,"#1e3a5f"],[0.4,"#3b82f6"],[1,"#6ee7b7"]],showscale=False),
            text=sd["sales"].apply(lambda v:f"${v:,.0f}"),textposition="outside"))
        fig4.update_layout(**dark_layout("Revenue by Sub-Category (Top 15)",420)); charts["subcat_bar"]=fig_to_json(fig4)
        if profit_col and "profit" in sd.columns:
            fig5=go.Figure()
            fig5.add_trace(go.Bar(name="Revenue",x=sd[subcat_col],y=sd["sales"],marker_color="#60a5fa"))
            fig5.add_trace(go.Bar(name="Profit",x=sd[subcat_col],y=sd["profit"],marker_color="#6ee7b7"))
            fig5.update_layout(**dark_layout("Sub-Category: Revenue vs Profit",320),barmode="group")
            charts["subcat_compare"]=fig_to_json(fig5)
    return {"charts":charts,"table":table_data}

def analyse_products(df):
    sales_col=safe_col(df,"sales","revenue","amount"); profit_col=safe_col(df,"profit")
    product_col=safe_col(df,"product_name","product"); qty_col=safe_col(df,"quantity","qty")
    charts={}
    if not sales_col or not product_col:
        return {"charts":charts,"table":[],"error":"Need product + sales columns"}
    agg_d={"sales":(sales_col,"sum"),"orders":(sales_col,"count")}
    if profit_col: agg_d["profit"]=(profit_col,"sum")
    if qty_col: agg_d["qty"]=(qty_col,"sum")
    prod=df.groupby(product_col).agg(**agg_d).reset_index()
    top20=prod.sort_values("sales",ascending=False).head(20)
    t15=top20.head(15)
    fig=go.Figure(go.Bar(x=t15["sales"],y=t15[product_col],orientation="h",
        marker=dict(color=t15["sales"],colorscale=[[0,"#1e3a5f"],[0.5,"#3b82f6"],[1,"#6ee7b7"]],showscale=False),
        text=t15["sales"].apply(lambda v:f"${v:,.0f}"),textposition="outside",
        hovertemplate="<b>%{y}</b><br>$%{x:,.0f}<extra></extra>"))
    fig.update_layout(**dark_layout("Top 15 Products by Revenue",480)); charts["top_products"]=fig_to_json(fig)
    if profit_col and "profit" in prod.columns:
        top50=prod.nlargest(50,"sales")
        fig2=go.Figure(go.Scatter(x=top50["sales"],y=top50["profit"],mode="markers",
            marker=dict(size=top50["sales"]/(top50["sales"].max()/25)+4,
                color=top50["profit"],colorscale=[[0,"#f87171"],[0.5,"#fbbf24"],[1,"#6ee7b7"]],showscale=True,
                colorbar=dict(title="Profit",tickfont=dict(color="#94a3b8",size=9))),
            text=top50[product_col],hovertemplate="<b>%{text}</b><br>Sales:$%{x:,.0f}<br>Profit:$%{y:,.0f}<extra></extra>"))
        fig2.update_layout(**dark_layout("Product: Sales vs Profit Bubble",380)); charts["bubble"]=fig_to_json(fig2)
        loss=prod[prod["profit"]<0].sort_values("profit").head(10)
        if len(loss):
            fig3=go.Figure(go.Bar(x=loss["profit"],y=loss[product_col],orientation="h",
                marker_color="#f87171",text=loss["profit"].apply(lambda v:f"${v:,.0f}"),textposition="outside"))
            fig3.update_layout(**dark_layout("Top Loss-Making Products",max(200,len(loss)*30+80)))
            charts["loss_products"]=fig_to_json(fig3)
    table=top20.rename(columns={product_col:"product"}).to_dict("records")
    return {"charts":charts,"table":table}

def analyse_customers(df):
    sales_col=safe_col(df,"sales","revenue","amount"); profit_col=safe_col(df,"profit")
    customer_col=safe_col(df,"customer_name","customer"); date_col=safe_col(df,"date","order_date")
    charts={}
    if not sales_col or not customer_col:
        return {"charts":charts,"table":[],"error":"Need customer + sales columns"}
    agg_d={"sales":(sales_col,"sum"),"orders":(sales_col,"count")}
    if profit_col: agg_d["profit"]=(profit_col,"sum")
    cust=df.groupby(customer_col).agg(**agg_d).reset_index()
    top20=cust.sort_values("sales",ascending=False).head(20)
    t15=top20.head(15)
    fig=go.Figure(go.Bar(x=t15["sales"],y=t15[customer_col],orientation="h",
        marker=dict(color=t15["sales"],colorscale=[[0,"#312e81"],[0.5,"#a78bfa"],[1,"#6ee7b7"]],showscale=False),
        text=t15["sales"].apply(lambda v:f"${v:,.0f}"),textposition="outside"))
    fig.update_layout(**dark_layout("Top 15 Customers by Revenue",460)); charts["top_customers"]=fig_to_json(fig)
    fig2=go.Figure(go.Histogram(x=cust["orders"],nbinsx=20,
        marker_color="rgba(96,165,250,0.7)",marker_line=dict(color="#1e3a5f",width=1)))
    fig2.update_layout(**dark_layout("Order Frequency Distribution",260)); charts["order_dist"]=fig_to_json(fig2)
    total=cust["sales"].sum()
    cust_s=cust.sort_values("sales",ascending=False).copy()
    cust_s["cum_pct"]=cust_s["sales"].cumsum()/total*100
    tA=cust_s[cust_s["cum_pct"]<=50].shape[0]
    tB=cust_s[(cust_s["cum_pct"]>50)&(cust_s["cum_pct"]<=80)].shape[0]
    tC=cust_s[cust_s["cum_pct"]>80].shape[0]
    fig3=go.Figure(go.Pie(
        labels=[f"Tier A ({tA})",f"Tier B ({tB})",f"Tier C ({tC})"],values=[tA,tB,tC],hole=0.55,
        marker=dict(colors=["#6ee7b7","#60a5fa","#475569"])))
    fig3.update_layout(**dark_layout("Customer Revenue Segmentation",290)); charts["segment_donut"]=fig_to_json(fig3)
    fig4=go.Figure(go.Scatter(x=cust["orders"],y=cust["sales"],mode="markers",
        marker=dict(size=6,opacity=0.65,color=cust["sales"],
            colorscale=[[0,"#1e3a5f"],[0.5,"#a78bfa"],[1,"#6ee7b7"]],showscale=False),
        text=cust[customer_col],hovertemplate="<b>%{text}</b><br>Orders:%{x}<br>Revenue:$%{y:,.0f}<extra></extra>"))
    fig4.update_layout(**dark_layout("Customer: Orders vs Revenue",280)); charts["scatter"]=fig_to_json(fig4)
    summary=dict(total_customers=int(cust.shape[0]),
                 avg_revenue_per_customer=round(float(cust["sales"].mean()),2),
                 top_customer=str(top20.iloc[0][customer_col]) if len(top20) else "—",
                 top_customer_revenue=round(float(top20.iloc[0]["sales"]),2) if len(top20) else 0)
    table=top20.rename(columns={customer_col:"customer"}).to_dict("records")
    return {"charts":charts,"table":table,"summary":summary}

def analyse_regional(df):
    sales_col=safe_col(df,"sales","revenue","amount"); profit_col=safe_col(df,"profit")
    region_col=safe_col(df,"region","state","country"); city_col=safe_col(df,"city")
    cat_col=safe_col(df,"category","product_category"); charts={}
    if not sales_col or not region_col:
        return {"charts":charts,"table":[],"error":"Need region + sales columns"}
    agg_d={"sales":(sales_col,"sum"),"orders":(sales_col,"count")}
    if profit_col: agg_d["profit"]=(profit_col,"sum")
    reg=df.groupby(region_col).agg(**agg_d).reset_index().sort_values("sales",ascending=False)
    fig=go.Figure()
    fig.add_trace(go.Bar(name="Revenue",x=reg[region_col],y=reg["sales"],marker_color="#60a5fa",
        text=reg["sales"].apply(lambda v:f"${v/1e3:.1f}K"),textposition="outside"))
    if profit_col:
        fig.add_trace(go.Bar(name="Profit",x=reg[region_col],y=reg["profit"],marker_color="#6ee7b7"))
    fig.update_layout(**dark_layout("Revenue & Profit by Region",300),barmode="group"); charts["region_bar"]=fig_to_json(fig)
    if cat_col and sales_col:
        rc=df.groupby([region_col,cat_col])[sales_col].sum().reset_index()
        rc.columns=["region","category","sales"]
        fig2=go.Figure(go.Treemap(labels=rc["category"],parents=rc["region"],values=rc["sales"],
            marker=dict(colorscale=[[0,"#0f1829"],[0.5,"#1d4ed8"],[1,"#6ee7b7"]]),
            hovertemplate="<b>%{label}</b><br>%{parent}<br>$%{value:,.0f}<extra></extra>"))
        fig2.update_layout(**dark_layout("Revenue by Region & Category",340)); charts["treemap"]=fig_to_json(fig2)
    if profit_col:
        reg["margin"]=(reg["profit"]/reg["sales"]*100).round(1)
        fig3=go.Figure(go.Bar(x=reg[region_col],y=reg["margin"],
            marker_color=["#6ee7b7" if v>0 else "#f87171" for v in reg["margin"]],
            text=reg["margin"].apply(lambda v:f"{v:.1f}%"),textposition="outside"))
        fig3.update_layout(**dark_layout("Profit Margin % by Region",260)); charts["margin_bar"]=fig_to_json(fig3)
    if city_col:
        city=df.groupby(city_col)[sales_col].sum().reset_index().sort_values(sales_col,ascending=False).head(15)
        fig4=go.Figure(go.Bar(x=city[sales_col],y=city[city_col],orientation="h",
            marker=dict(color=city[sales_col],colorscale=[[0,"#1e3a5f"],[0.5,"#f472b6"],[1,"#fbbf24"]],showscale=False),
            text=city[sales_col].apply(lambda v:f"${v:,.0f}"),textposition="outside"))
        fig4.update_layout(**dark_layout("Top 15 Cities by Revenue",440)); charts["city_bar"]=fig_to_json(fig4)
    table=reg.rename(columns={region_col:"region"}).to_dict("records")
    return {"charts":charts,"table":table}

def analyse_insights_full(df):
    sales_col=safe_col(df,"sales","revenue","amount"); profit_col=safe_col(df,"profit")
    cat_col=safe_col(df,"category","product_category"); region_col=safe_col(df,"region","state","country")
    customer_col=safe_col(df,"customer_name","customer"); product_col=safe_col(df,"product_name","product")
    date_col=safe_col(df,"date","order_date"); qty_col=safe_col(df,"quantity","qty")
    ts=float(df[sales_col].sum()) if sales_col else 0
    tp=float(df[profit_col].sum()) if profit_col else 0
    kpis=dict(total_sales=ts,total_profit=tp,total_orders=int(df.shape[0]),
              profit_margin=tp/ts*100 if ts else 0,avg_order_value=ts/df.shape[0] if df.shape[0] else 0,
              unique_customers=int(df[customer_col].nunique()) if customer_col else 0,
              total_quantity=int(df[qty_col].sum()) if qty_col else 0,
              loss_orders=int((df[profit_col]<0).sum()) if profit_col else 0)
    ins=_insights(df,kpis,sales_col,profit_col,cat_col,region_col,customer_col,qty_col)
    if date_col and sales_col:
        grp=df.dropna(subset=[date_col]).copy(); grp["_m"]=grp[date_col].dt.to_period("M")
        mo=grp.groupby("_m")[sales_col].sum().reset_index()
        if len(mo)>=2:
            last=float(mo.iloc[-1][sales_col]); prev=float(mo.iloc[-2][sales_col])
            chg=(last-prev)/prev*100 if prev else 0
            t="positive" if chg>=0 else "negative"
            ins.append({"type":t,"icon":"📈" if chg>=0 else "📉","text":f"Most recent month changed {chg:+.1f}% vs prior month (${last:,.0f} vs ${prev:,.0f})."})
    if product_col and profit_col and sales_col:
        prod=df.groupby(product_col).agg(sales=(sales_col,"sum"),profit=(profit_col,"sum")).reset_index()
        best=prod.loc[prod["profit"].idxmax()]
        ins.append({"type":"positive","icon":"⭐","text":f"Most profitable product: '{best[product_col]}' — ${best['profit']:,.0f} profit on ${best['sales']:,.0f} revenue."})
    if cat_col and profit_col and sales_col:
        cm=df.groupby(cat_col).agg(sales=(sales_col,"sum"),profit=(profit_col,"sum")).reset_index()
        cm["margin"]=cm["profit"]/cm["sales"]*100
        worst=cm.loc[cm["margin"].idxmin()]
        ins.append({"type":"warning","icon":"🔍","text":f"Category '{worst[cat_col]}' has lowest margin at {worst['margin']:.1f}%. Review pricing or cost structures."})
    return {"insights":ins,"kpis":kpis}

def build_excel(df):
    sales_col=safe_col(df,"sales","revenue","amount"); profit_col=safe_col(df,"profit")
    cat_col=safe_col(df,"category","product_category"); product_col=safe_col(df,"product_name","product")
    customer_col=safe_col(df,"customer_name","customer"); region_col=safe_col(df,"region","state","country")
    date_col=safe_col(df,"date","order_date"); qty_col=safe_col(df,"quantity","qty")
    ts=float(df[sales_col].sum()) if sales_col else 0
    tp=float(df[profit_col].sum()) if profit_col else 0
    to=int(df.shape[0]); aov=ts/to if to else 0; margin=tp/ts*100 if ts else 0
    uc=int(df[customer_col].nunique()) if customer_col else 0
    tq=int(df[qty_col].sum()) if qty_col else 0
    wb=openpyxl.Workbook()
    def sh(n): return Font(bold=True,color="FFFFFF",name="Calibri",size=n)
    def sb(): return Font(color="1E293B",name="Calibri",size=10)
    def sm(neg=False): return Font(color="991B1B" if neg else "065F46",name="Calibri",size=10,bold=True)
    F_D=PatternFill("solid",fgColor="0F172A"); F_G=PatternFill("solid",fgColor="059669")
    F_B=PatternFill("solid",fgColor="1D4ED8"); F_P=PatternFill("solid",fgColor="7C3AED")
    F_O=PatternFill("solid",fgColor="0369A1"); F_H=PatternFill("solid",fgColor="1E293B")
    F_A=PatternFill("solid",fgColor="F0FDF4"); F_W=PatternFill("solid",fgColor="FFFFFF")
    CTR=Alignment(horizontal="center",vertical="center",wrap_text=True)
    LFT=Alignment(horizontal="left",vertical="center",wrap_text=True)
    thin=Side(style="thin",color="E2E8F0"); BDR=Border(left=thin,right=thin,top=thin,bottom=thin)
    def hdr(ws,r,vals,fill,fnt,aln=None):
        for c,v in enumerate(vals,1):
            cell=ws.cell(row=r,column=c,value=v); cell.fill=fill; cell.font=fnt
            cell.alignment=aln or LFT; cell.border=BDR
    def brow(ws,r,vals,alt=False,mc=None):
        f=F_A if alt else F_W
        for c,v in enumerate(vals,1):
            cell=ws.cell(row=r,column=c,value=v); cell.fill=f; cell.border=BDR
            if mc and c in mc: cell.font=sm(isinstance(v,float) and v<0)
            else: cell.font=sb()
            cell.alignment=LFT
    def aw(ws,ex=4,cap=50):
        for col in ws.columns:
            w=max((len(str(c.value or "")) for c in col),default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width=min(w+ex,cap)
    def sec(ws,r,txt,nc=2):
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=nc)
        c=ws.cell(row=r,column=1,value=txt); c.fill=F_G; c.font=sh(11); c.alignment=CTR
        ws.row_dimensions[r].height=20
    # Sheet 1
    ws1=wb.active; ws1.title="📊 Executive Summary"
    ws1.merge_cells("A1:E1"); ws1["A1"]="NexusIQ — E-Commerce Analytics Report"
    ws1["A1"].font=sh(16); ws1["A1"].fill=F_D; ws1["A1"].alignment=CTR; ws1.row_dimensions[1].height=36
    ws1.merge_cells("A2:E2")
    ws1["A2"]=f"Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}  |  {df.shape[0]:,} rows × {df.shape[1]} cols"
    ws1["A2"].font=Font(color="64748B",name="Calibri",size=9); ws1["A2"].fill=F_D; ws1["A2"].alignment=CTR; ws1.row_dimensions[2].height=16
    ws1.append([])
    sec(ws1,4,"KEY PERFORMANCE INDICATORS",2)
    hdr(ws1,5,["Metric","Value"],F_H,sh(11),CTR)
    kd=[("Total Revenue",f"${ts:,.2f}"),("Total Profit",f"${tp:,.2f}"),("Profit Margin",f"{margin:.2f}%"),
        ("Total Orders",f"{to:,}"),("Total Units Sold",f"{tq:,}"),("Avg Order Value",f"${aov:,.2f}"),("Unique Customers",f"{uc:,}")]
    for i,(k,v) in enumerate(kd,6): brow(ws1,i,[k,v],alt=i%2==0)
    r=len(kd)+8; ws1.append([]); sec(ws1,r,"AI BUSINESS INSIGHTS",2); r+=1
    hdr(ws1,r,["Type","Insight"],F_H,sh(11),CTR); r+=1
    ov=analyse_overview(df)
    for i,ins in enumerate(ov["insights"]): brow(ws1,r+i,[ins["icon"]+" "+ins["type"].upper(),ins["text"]],alt=i%2==0)
    aw(ws1)
    # Sheet 2 Monthly
    if date_col and sales_col:
        ws2=wb.create_sheet("📈 Monthly Trends")
        grp=df.dropna(subset=[date_col]).copy(); grp["Month"]=grp[date_col].dt.to_period("M").astype(str)
        ad={"Revenue":(sales_col,"sum"),"Orders":(sales_col,"count")}
        if profit_col: ad["Profit"]=(profit_col,"sum")
        if qty_col: ad["Quantity"]=(qty_col,"sum")
        m=grp.groupby("Month").agg(**ad).reset_index()
        if profit_col and "Profit" in m: m["Margin%"]=(m["Profit"]/m["Revenue"]*100).round(2)
        m["MoM Growth%"]=m["Revenue"].pct_change()*100
        hdr(ws2,1,list(m.columns),F_H,sh(11),CTR)
        for i,rd in enumerate(m.values.tolist(),2):
            brow(ws2,i,[round(v,2) if isinstance(v,float) else v for v in rd],alt=i%2==0,mc={2,4})
        aw(ws2)
    # Sheet 3 Categories
    if cat_col and sales_col:
        ws3=wb.create_sheet("🧩 Categories")
        ad={"Revenue":(sales_col,"sum"),"Orders":(sales_col,"count")}
        if profit_col: ad["Profit"]=(profit_col,"sum")
        if qty_col: ad["Quantity"]=(qty_col,"sum")
        cd=df.groupby(cat_col).agg(**ad).reset_index(); cd.columns=["Category"]+list(cd.columns[1:])
        cd=cd.sort_values("Revenue",ascending=False)
        if profit_col and "Profit" in cd.columns: cd["Margin%"]=(cd["Profit"]/cd["Revenue"]*100).round(2)
        cd["Revenue Share%"]=(cd["Revenue"]/cd["Revenue"].sum()*100).round(2)
        hdr(ws3,1,list(cd.columns),F_G,sh(11),CTR)
        for i,rd in enumerate(cd.values.tolist(),2):
            brow(ws3,i,[round(v,2) if isinstance(v,float) else v for v in rd],alt=i%2==0,mc={2,3})
        aw(ws3)
    # Sheet 4 Products
    if product_col and sales_col:
        ws4=wb.create_sheet("📦 Top Products")
        ad={"Revenue":(sales_col,"sum"),"Orders":(sales_col,"count")}
        if profit_col: ad["Profit"]=(profit_col,"sum")
        if qty_col: ad["Qty Sold"]=(qty_col,"sum")
        prod=df.groupby(product_col).agg(**ad).reset_index(); prod.columns=["Product"]+list(prod.columns[1:])
        prod=prod.sort_values("Revenue",ascending=False).head(50)
        if profit_col and "Profit" in prod.columns: prod["Margin%"]=(prod["Profit"]/prod["Revenue"]*100).round(2)
        prod.insert(0,"Rank",range(1,len(prod)+1))
        hdr(ws4,1,list(prod.columns),F_B,sh(11),CTR)
        for i,rd in enumerate(prod.values.tolist(),2):
            brow(ws4,i,[round(v,2) if isinstance(v,float) else v for v in rd],alt=i%2==0,mc={3,4})
        aw(ws4)
    # Sheet 5 Customers
    if customer_col and sales_col:
        ws5=wb.create_sheet("👤 Top Customers")
        ad={"Revenue":(sales_col,"sum"),"Orders":(sales_col,"count")}
        if profit_col: ad["Profit"]=(profit_col,"sum")
        cust=df.groupby(customer_col).agg(**ad).reset_index(); cust.columns=["Customer"]+list(cust.columns[1:])
        cust=cust.sort_values("Revenue",ascending=False).head(50)
        cust["Avg Order Value"]=(cust["Revenue"]/cust["Orders"]).round(2)
        cust.insert(0,"Rank",range(1,len(cust)+1))
        hdr(ws5,1,list(cust.columns),F_P,sh(11),CTR)
        for i,rd in enumerate(cust.values.tolist(),2):
            brow(ws5,i,[round(v,2) if isinstance(v,float) else v for v in rd],alt=i%2==0,mc={3,4})
        aw(ws5)
    # Sheet 6 Regional
    if region_col and sales_col:
        ws6=wb.create_sheet("🌍 Regional")
        ad={"Revenue":(sales_col,"sum"),"Orders":(sales_col,"count")}
        if profit_col: ad["Profit"]=(profit_col,"sum")
        reg=df.groupby(region_col).agg(**ad).reset_index(); reg.columns=["Region"]+list(reg.columns[1:])
        reg=reg.sort_values("Revenue",ascending=False)
        if profit_col and "Profit" in reg.columns: reg["Margin%"]=(reg["Profit"]/reg["Revenue"]*100).round(2)
        reg["Revenue Share%"]=(reg["Revenue"]/reg["Revenue"].sum()*100).round(2)
        hdr(ws6,1,list(reg.columns),F_O,sh(11),CTR)
        for i,rd in enumerate(reg.values.tolist(),2):
            brow(ws6,i,[round(v,2) if isinstance(v,float) else v for v in rd],alt=i%2==0,mc={2,3})
        aw(ws6)
    # Sheet 7 Raw
    ws7=wb.create_sheet("📋 Raw Data")
    hdr(ws7,1,df.columns.tolist(),F_D,sh(11),CTR)
    for i,rd in enumerate(df.head(5000).values.tolist(),2):
        for c,v in enumerate(rd,1):
            cell=ws7.cell(row=i,column=c,value=v if not pd.isna(v) else "")
            cell.font=sb(); cell.alignment=LFT
    aw(ws7,ex=2,cap=40)
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

@app.route("/")
def index():
    with open("index.html","r") as f: return f.read()

@app.route("/upload",methods=["POST"])
def upload():
    if "file" not in request.files: return jsonify({"error":"No file provided"}),400
    file=request.files["file"]
    if not file.filename: return jsonify({"error":"Empty filename"}),400
    fname=file.filename.lower()
    try:
        if fname.endswith(".csv"): df=pd.read_csv(file,encoding="utf-8",on_bad_lines="skip")
        elif fname.endswith((".xlsx",".xls")): df=pd.read_excel(file)
        else: return jsonify({"error":"Unsupported file type. Upload CSV or Excel."}),400
        df=clean_dataframe(df); df.to_pickle(os.path.join(UPLOAD_FOLDER,"current.pkl"))
        return jsonify({"success":True,"data":analyse_overview(df)})
    except Exception as e:
        traceback.print_exc(); return jsonify({"error":str(e)}),500

@app.route("/api/trends")
def api_trends():
    df=load_df();
    if df is None: return jsonify({"error":"No dataset loaded"}),400
    return jsonify(analyse_trends(df))

@app.route("/api/categories")
def api_categories():
    df=load_df()
    if df is None: return jsonify({"error":"No dataset loaded"}),400
    return jsonify(analyse_categories(df))

@app.route("/api/products")
def api_products():
    df=load_df()
    if df is None: return jsonify({"error":"No dataset loaded"}),400
    return jsonify(analyse_products(df))

@app.route("/api/customers")
def api_customers():
    df=load_df()
    if df is None: return jsonify({"error":"No dataset loaded"}),400
    return jsonify(analyse_customers(df))

@app.route("/api/regional")
def api_regional():
    df=load_df()
    if df is None: return jsonify({"error":"No dataset loaded"}),400
    return jsonify(analyse_regional(df))

@app.route("/api/insights")
def api_insights():
    df=load_df()
    if df is None: return jsonify({"error":"No dataset loaded"}),400
    return jsonify(analyse_insights_full(df))

@app.route("/report/excel")
def export_excel():
    df=load_df()
    if df is None: return jsonify({"error":"No dataset loaded. Upload a file first."}),400
    try:
        df=clean_dataframe(df); buf=build_excel(df)
        return send_file(buf,as_attachment=True,
            download_name=f"nexusiq_report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        traceback.print_exc(); return jsonify({"error":str(e)}),500

if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0",port=port,debug=False)
