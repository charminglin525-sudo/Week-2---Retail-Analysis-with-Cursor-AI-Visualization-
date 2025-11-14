import pandas as pd
import numpy as np
import warnings
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime

# 抑制 Streamlit 的 ScriptRunContext 警告（在 bare mode 下可以安全忽略）
warnings.filterwarnings("ignore", message=".*missing ScriptRunContext.*")

# 設置頁面配置
st.set_page_config(
    page_title="E-commerce Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定義CSS樣式
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .kpi-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    .insight-box {
        background-color: #fff3cd;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #ffc107;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# 查找列名的函數
def find_column(df, keywords):
    """查找包含關鍵詞的列名"""
    if df is None or len(df) == 0 or len(df.columns) == 0:
        return None
    keywords_lower = [k.lower() for k in keywords]
    for col in df.columns:
        col_lower = col.lower()
        for keyword in keywords_lower:
            if keyword in col_lower:
                return col
    return None

# 加載數據
@st.cache_data
def load_data():
    """加載並處理數據"""
    try:
        # 從彙總表.xlsx讀取數據
        # 讀取MOM數據
        try:
            mom_df = pd.read_excel('彙總表.xlsx', sheet_name='MOM')
            st.success(f"✓ 成功加載 MOM 數據: {len(mom_df)} 行")
        except Exception as e:
            st.error(f"✗ 無法加載 MOM 數據: {e}")
            mom_df = pd.DataFrame()
        
        # 讀取AOV_ARPU數據
        try:
            aov_arpu_df = pd.read_excel('彙總表.xlsx', sheet_name='AOV_ARPU')
            st.success(f"✓ 成功加載 AOV_ARPU 數據: {len(aov_arpu_df)} 行")
        except Exception as e:
            st.warning(f"⚠ 無法加載 AOV_ARPU 數據: {e}")
            aov_arpu_df = pd.DataFrame()
        
        # 讀取RFM數據
        try:
            rfm_df = pd.read_excel('彙總表.xlsx', sheet_name='RFM')
            st.success(f"✓ 成功加載 RFM 數據: {len(rfm_df)} 行")
        except Exception as e:
            st.warning(f"⚠ 無法加載 RFM 數據: {e}")
            rfm_df = pd.DataFrame()
        
        # 讀取SKU數據
        try:
            sku_df = pd.read_excel('彙總表.xlsx', sheet_name='SKU')
            st.success(f"✓ 成功加載 SKU 數據: {len(sku_df)} 行")
        except Exception as e:
            st.warning(f"⚠ 無法加載 SKU 數據: {e}")
            sku_df = pd.DataFrame()
        
        # 讀取Sales by Country數據
        try:
            sales_by_country_df = pd.read_excel('彙總表.xlsx', sheet_name='Sales by Country')
            st.success(f"✓ 成功加載 Sales by Country 數據: {len(sales_by_country_df)} 行")
        except Exception as e:
            st.warning(f"⚠ 無法加載 Sales by Country 數據: {e}")
            sales_by_country_df = pd.DataFrame()
        
        # 讀取Return and Abnormal數據（如果存在）
        return_product_df = pd.DataFrame()
        abnormal_product_df = pd.DataFrame()
        try:
            return_abnormal_file = 'Return and Abnormal_2011_11.xlsx'
            return_product_df = pd.read_excel(return_abnormal_file, sheet_name='Return analysis product')
            abnormal_product_df = pd.read_excel(return_abnormal_file, sheet_name='Abnormal analysis product')
            st.success(f"✓ 成功加載 Return and Abnormal 數據")
        except:
            # 嘗試其他可能的文件名
            try:
                return_abnormal_file = 'Return and Abnormal.xlsx'
                return_product_df = pd.read_excel(return_abnormal_file, sheet_name='Return analysis product')
                abnormal_product_df = pd.read_excel(return_abnormal_file, sheet_name='Abnormal analysis product')
                st.success(f"✓ 成功加載 Return and Abnormal 數據")
            except:
                st.info("ℹ 未找到 Return and Abnormal 數據文件（可選）")
        
        return {
            'mom': mom_df,
            'aov_arpu': aov_arpu_df,
            'rfm': rfm_df,
            'sku': sku_df,
            'sales_by_country': sales_by_country_df,
            'return_product': return_product_df,
            'abnormal_product': abnormal_product_df
        }
    except Exception as e:
        st.error(f"加載數據時發生嚴重錯誤: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

# 篩選2011/01 - 2011/11的數據
def filter_2011_data(df, date_column='YearMonth'):
    """篩選2011年1月到11月的數據"""
    if df is None or len(df) == 0:
        return df
    
    if date_column in df.columns:
        # 轉換YearMonth為字符串格式以便篩選
        df[date_column] = df[date_column].astype(str)
        # 篩選2011年的數據
        filtered = df[df[date_column].str.startswith('2011')].copy()
        # 排除2011年12月
        filtered = filtered[~filtered[date_column].str.contains('2011-12')]
        return filtered
    return df

# 生成KPI卡片
def generate_kpi(data):
    """生成KPI概覽卡片（只顯示最後一個月的數據）"""
    if data is None:
        st.error("無法加載數據")
        return
    
    mom_df = filter_2011_data(data['mom'])
    aov_arpu_df = filter_2011_data(data['aov_arpu'])
    
    if len(mom_df) == 0:
        st.warning("沒有2011年的數據")
        return
    
    # 獲取最後一個月的數據
    last_month_data = mom_df.iloc[-1]
    last_month = last_month_data['YearMonth'] if 'YearMonth' in last_month_data.index else 'N/A'
    
    # 解析年月（格式可能是 "2011-11" 或 "2011/11"）
    try:
        if '-' in str(last_month):
            year, month = str(last_month).split('-')
        elif '/' in str(last_month):
            year, month = str(last_month).split('/')
        else:
            year = str(last_month)[:4]
            month = str(last_month)[4:6] if len(str(last_month)) >= 6 else 'N/A'
        
        month_name = f"{year}年{month}月"
    except:
        month_name = str(last_month)
    
    # 顯示標題和月份信息
    st.markdown(f'<div class="main-header">📊 E-commerce Dashboard - {month_name}</div>', unsafe_allow_html=True)
    st.markdown(f"**數據期間: {month_name}**")
    
    # 使用最後一個月的數據計算KPI
    last_revenue = last_month_data['Revenue'] if 'Revenue' in last_month_data.index else 0
    last_orders = last_month_data['Normal_Orders'] if 'Normal_Orders' in last_month_data.index else 0
    last_customers = last_month_data['Customer'] if 'Customer' in last_month_data.index else 0
    last_return_orders = last_month_data['Return_Orders'] if 'Return_Orders' in last_month_data.index else 0
    last_return_amount = abs(last_month_data['Return']) if 'Return' in last_month_data.index else 0
    
    # 計算退貨率
    return_rate = (last_return_orders / (last_return_orders + last_orders) * 100) if (last_return_orders + last_orders) > 0 else 0
    
    # 計算AOV
    aov = (last_revenue / last_orders) if last_orders > 0 else 0
    
    # 計算ARPU（從AOV_ARPU數據）
    last_arpu = 0
    if len(aov_arpu_df) > 0:
        # 找到最後一個月對應的ARPU
        last_month_str = str(last_month)
        aov_arpu_filtered = aov_arpu_df[aov_arpu_df['YearMonth'] == last_month_str]
        if len(aov_arpu_filtered) > 0:
            last_arpu = aov_arpu_filtered.iloc[0]['ARPU'] if 'ARPU' in aov_arpu_filtered.columns else 0
        else:
            # 如果找不到，使用計算值
            last_arpu = (last_revenue / last_customers) if last_customers > 0 else 0
    
    # 計算MoM增長率（最後一個月相對於前一個月）
    if len(mom_df) >= 2:
        prev_month_data = mom_df.iloc[-2]
        prev_revenue = prev_month_data['Revenue'] if 'Revenue' in prev_month_data.index else 0
        revenue_mom = ((last_revenue - prev_revenue) / prev_revenue * 100) if prev_revenue > 0 else 0
    else:
        revenue_mom = 0
    
    # 第一行KPI卡片：Revenue, Orders, Customers
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            label="Revenue",
            value=f"${last_revenue:,.0f}",
            delta=f"{revenue_mom:.1f}% MoM" if revenue_mom != 0 else None
        )
    
    with col2:
        st.metric(
            label="Orders",
            value=f"{last_orders:,.0f}",
            delta=None
        )
    
    with col3:
        st.metric(
            label="Customers",
            value=f"{last_customers:,.0f}",
            delta=None
        )
    
    # 第二行KPI卡片：AOV, ARPU
    col4, col5 = st.columns(2)
    
    with col4:
        st.metric(
            label="AOV",
            value=f"${aov:.2f}",
            delta=None
        )
    
    with col5:
        st.metric(
            label="ARPU",
            value=f"${last_arpu:.2f}",
            delta=None
        )
    
    # 第三行KPI卡片：Return Amount, Return Orders, Return Rate
    col6, col7, col8 = st.columns(3)
    
    with col6:
        st.metric(
            label="Return Amount",
            value=f"${last_return_amount:,.0f}",
            delta=None
        )
    
    with col7:
        st.metric(
            label="Return Orders",
            value=f"{last_return_orders:,.0f}",
            delta=None
        )
    
    with col8:
        st.metric(
            label="Return Rate",
            value=f"{return_rate:.2f}%",
            delta=None
        )

# 生成月度趨勢圖表
def generate_mom_charts(data):
    """生成月度趨勢圖表（包含 Revenue, Orders, Customer, AOV, ARPU）"""
    st.markdown("## 📈 Monthly Trends (MOM)")
    
    if data is None:
        return
    
    mom_df = filter_2011_data(data['mom'])
    aov_arpu_df = filter_2011_data(data['aov_arpu'])
    
    if len(mom_df) == 0:
        st.warning("沒有2011年的數據")
        return
    
    # 第一張圖：Revenue 和 Orders 線圖（使用雙Y軸）
    fig1 = make_subplots(specs=[[{"secondary_y": True}]])
    
    if 'Revenue' in mom_df.columns:
        fig1.add_trace(
            go.Scatter(
                x=mom_df['YearMonth'],
                y=mom_df['Revenue'],
                name='Revenue',
                line=dict(color='#1f77b4', width=3),
                mode='lines+markers'
            ),
            row=1, col=1, secondary_y=False
        )
    
    if 'Normal_Orders' in mom_df.columns:
        fig1.add_trace(
            go.Scatter(
                x=mom_df['YearMonth'],
                y=mom_df['Normal_Orders'],
                name='Orders',
                line=dict(color='#ff7f0e', width=3),
                mode='lines+markers'
            ),
            row=1, col=1, secondary_y=True
        )
    
    # 標記負增長月份
    if 'Revenue_Growth' in mom_df.columns:
        negative_months = mom_df[mom_df['Revenue_Growth'] < 0]
        if len(negative_months) > 0:
            fig1.add_trace(
                go.Scatter(
                    x=negative_months['YearMonth'],
                    y=negative_months['Revenue'],
                    mode='markers',
                    marker=dict(
                        symbol='x',
                        size=15,
                        color='red',
                        line=dict(width=2, color='red')
                    ),
                    name='Negative Growth',
                    showlegend=True
                ),
                row=1, col=1, secondary_y=False
            )
    
    fig1.update_xaxes(title_text="Month", row=1, col=1)
    fig1.update_yaxes(title_text="Revenue ($)", row=1, col=1, secondary_y=False)
    fig1.update_yaxes(title_text="Orders", row=1, col=1, secondary_y=True)
    fig1.update_layout(
        title="Revenue & Orders Trends",
        height=400,
        showlegend=True,
        hovermode='x unified'
    )
    st.plotly_chart(fig1, use_container_width=True)
    
    # 第二張圖：Customers 柱狀圖
    if 'Customer' in mom_df.columns:
        fig2 = go.Figure()
        fig2.add_trace(
            go.Bar(
                x=mom_df['YearMonth'],
                y=mom_df['Customer'],
                name='Customers',
                marker=dict(color='#2ca02c')
            )
        )
        fig2.update_xaxes(title_text="Month")
        fig2.update_yaxes(title_text="Customers")
        fig2.update_layout(
            title="Customers Trend",
            height=400,
            showlegend=False
        )
        st.plotly_chart(fig2, use_container_width=True)
    
    # 第三部分：AOV 和 ARPU 分開顯示（左右並排）
    col1, col2 = st.columns(2)
    
    # 合併數據以確保月份對齊
    if len(aov_arpu_df) > 0 and 'YearMonth' in aov_arpu_df.columns:
        merged_df = mom_df.merge(aov_arpu_df, on='YearMonth', how='left')
        
        with col1:
            # 左邊：AOV 線圖
            if 'AOV' in merged_df.columns:
                fig_aov = go.Figure()
                fig_aov.add_trace(
                    go.Scatter(
                        x=merged_df['YearMonth'],
                        y=merged_df['AOV'],
                        name='AOV',
                        line=dict(color='#9467bd', width=3),
                        mode='lines+markers'
                    )
                )
                fig_aov.update_xaxes(title_text="Month")
                fig_aov.update_yaxes(title_text="AOV ($)")
                fig_aov.update_layout(
                    title="AOV Trend",
                    height=400,
                    showlegend=False
                )
                st.plotly_chart(fig_aov, use_container_width=True)
        
        with col2:
            # 右邊：ARPU 線圖
            if 'ARPU' in merged_df.columns:
                fig_arpu = go.Figure()
                fig_arpu.add_trace(
                    go.Scatter(
                        x=merged_df['YearMonth'],
                        y=merged_df['ARPU'],
                        name='ARPU',
                        line=dict(color='#8c564b', width=3),
                        mode='lines+markers'
                    )
                )
                fig_arpu.update_xaxes(title_text="Month")
                fig_arpu.update_yaxes(title_text="ARPU ($)")
                fig_arpu.update_layout(
                    title="ARPU Trend",
                    height=400,
                    showlegend=False
                )
                st.plotly_chart(fig_arpu, use_container_width=True)

# 生成RFM可視化
def generate_rfm_visualization(data):
    """生成RFM客戶細分可視化"""
    st.markdown("## 👥 RFM Customer Segmentation")
    
    if data is None:
        return
    
    rfm_df = data['rfm']
    
    if len(rfm_df) == 0:
        st.warning("沒有RFM數據")
        return
    
    # 查找CustomerID列
    customer_id_col = find_column(rfm_df, ['CustomerID', 'Customer ID', 'Customer', 'customer'])
    
    # GUEST vs Others 比較
    if customer_id_col and 'Monetary' in rfm_df.columns:
        # 標識GUEST客戶
        rfm_df['IsGuest'] = rfm_df[customer_id_col].astype(str).str.strip().str.upper() == 'GUEST'
        
        # 計算GUEST vs Others的Monetary和Count
        guest_stats = rfm_df.groupby('IsGuest').agg({
            'Monetary': 'sum',
            customer_id_col: 'count'
        }).reset_index()
        guest_stats.columns = ['IsGuest', 'Monetary', 'Count']
        guest_stats['Type'] = guest_stats['IsGuest'].map({True: 'GUEST', False: 'Others'})
        
        # 顯示GUEST vs Others比較
        st.markdown("### GUEST vs Others Comparison")
        col1, col2 = st.columns(2)
        
        with col1:
            # Monetary比較
            fig_guest_monetary = px.bar(
                guest_stats,
                x='Type',
                y='Monetary',
                title='Monetary: GUEST vs Others',
                color='Type',
                color_discrete_map={'GUEST': '#e74c3c', 'Others': '#3498db'}
            )
            fig_guest_monetary.update_layout(height=400, showlegend=False)
            st.plotly_chart(fig_guest_monetary, use_container_width=True)
        
        with col2:
            # Count比較
            fig_guest_count = px.bar(
                guest_stats,
                x='Type',
                y='Count',
                title='Count: GUEST vs Others',
                color='Type',
                color_discrete_map={'GUEST': '#e74c3c', 'Others': '#3498db'}
            )
            fig_guest_count.update_layout(height=400, showlegend=False)
            st.plotly_chart(fig_guest_count, use_container_width=True)
        
        # 去掉GUEST進行後續分析
        rfm_df_no_guest = rfm_df[~rfm_df['IsGuest']].copy()
        guest_count = len(rfm_df) - len(rfm_df_no_guest)
        st.info(f"ℹ️ 已排除 {guest_count} 個GUEST客戶，以下分析僅包含註冊客戶")
    else:
        rfm_df_no_guest = rfm_df.copy()
        st.warning("⚠️ 無法識別GUEST客戶，將使用全部數據")
    
    # RFM Scatter Plot (Total Score vs Revenue, color = Category)
    if 'Total_Score' in rfm_df_no_guest.columns and 'Monetary' in rfm_df_no_guest.columns and 'Category' in rfm_df_no_guest.columns:
        st.markdown("### RFM Scatter Plot (Total Score vs Revenue)")
        # 確保Total_Score和Monetary是數值類型
        rfm_scatter_df = rfm_df_no_guest.copy()
        rfm_scatter_df['Total_Score'] = pd.to_numeric(rfm_scatter_df['Total_Score'], errors='coerce')
        rfm_scatter_df['Monetary'] = pd.to_numeric(rfm_scatter_df['Monetary'], errors='coerce')
        rfm_scatter_df = rfm_scatter_df.dropna(subset=['Total_Score', 'Monetary', 'Category'])
        
        if len(rfm_scatter_df) > 0:
            # 定義顏色映射（從Champions到Lost：深藍色到深紅色）
            category_order = ['Champions', 'Loyal', 'Potential Loyalist', 'At Risk', 'Lost', 'Unknown']
            colors_blue_to_red = ['#1a237e', '#3949ab', '#5c6bc0', '#e64a19', '#c62828', '#95a5a6']
            color_map = dict(zip(category_order, colors_blue_to_red))
            
            # 創建散點圖
            # 準備hover_data
            hover_data_list = []
            if customer_id_col and customer_id_col in rfm_scatter_df.columns:
                hover_data_list.append(customer_id_col)
            if 'Frequency' in rfm_scatter_df.columns:
                hover_data_list.append('Frequency')
            if 'Recency' in rfm_scatter_df.columns:
                hover_data_list.append('Recency')
            
            fig_scatter = px.scatter(
                rfm_scatter_df,
                x='Total_Score',
                y='Monetary',
                color='Category',
                title='RFM Scatter Plot (Total Score vs Revenue)',
                labels={
                    'Total_Score': 'Total Score',
                    'Monetary': 'Revenue (Monetary)',
                    'Category': 'Category'
                },
                color_discrete_map=color_map,
                hover_data=hover_data_list if hover_data_list else None
            )
            fig_scatter.update_layout(
                height=500,
                xaxis_title='Total Score',
                yaxis_title='Revenue (Monetary)',
                showlegend=True
            )
            st.plotly_chart(fig_scatter, use_container_width=True)
        else:
            st.warning("無法創建散點圖：Total_Score和Monetary必須是數值類型")
    
    # Revenue Contribution和Customer Contribution (Pie Charts)
    if 'Category' in rfm_df_no_guest.columns and 'Monetary' in rfm_df_no_guest.columns:
        # 定義顏色映射（從Champions到Lost：深藍色到深紅色）
        category_order = ['Champions', 'Loyal', 'Potential Loyalist', 'At Risk', 'Lost', 'Unknown']
        colors_blue_to_red = ['#1a237e', '#3949ab', '#5c6bc0', '#e64a19', '#c62828', '#95a5a6']
        color_map = dict(zip(category_order, colors_blue_to_red))
        
        # 計算各組的Revenue和Count
        # 使用第一列作為計數列（通常是CustomerID或索引）
        count_col = customer_id_col if customer_id_col else rfm_df_no_guest.columns[0]
        category_stats = rfm_df_no_guest.groupby('Category').agg({
            'Monetary': 'sum',
            count_col: 'count'
        }).reset_index()
        category_stats.columns = ['Category', 'Revenue', 'Count']
        
        # 確保Category按照定義的順序
        category_stats['Category'] = pd.Categorical(category_stats['Category'], categories=category_order, ordered=True)
        category_stats = category_stats.sort_values('Category')
        
        # 計算占比
        total_revenue = category_stats['Revenue'].sum()
        total_count = category_stats['Count'].sum()
        category_stats['Revenue_Pct'] = (category_stats['Revenue'] / total_revenue * 100).round(2)
        category_stats['Count_Pct'] = (category_stats['Count'] / total_count * 100).round(2)
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Revenue Contribution Pie Chart
            fig_revenue_pie = go.Figure(data=[go.Pie(
                labels=category_stats['Category'],
                values=category_stats['Revenue'],
                hole=0.3,
                marker=dict(colors=[color_map.get(cat, '#95a5a6') for cat in category_stats['Category']]),
                textinfo='label+percent',
                hovertemplate='<b>%{label}</b><br>Revenue: $%{value:,.0f}<br>Percentage: %{percent}<extra></extra>'
            )])
            fig_revenue_pie.update_layout(
                title='Revenue Contribution by RFM Category',
                height=500
            )
            st.plotly_chart(fig_revenue_pie, use_container_width=True)
            
            # 顯示詳細占比
            st.markdown("**Revenue占比：**")
            for _, row in category_stats.iterrows():
                st.write(f"- {row['Category']}: ${row['Revenue']:,.0f} ({row['Revenue_Pct']:.2f}%)")
        
        with col2:
            # Customer Contribution Pie Chart
            fig_count_pie = go.Figure(data=[go.Pie(
                labels=category_stats['Category'],
                values=category_stats['Count'],
                hole=0.3,
                marker=dict(colors=[color_map.get(cat, '#95a5a6') for cat in category_stats['Category']]),
                textinfo='label+percent',
                hovertemplate='<b>%{label}</b><br>Count: %{value:,.0f}<br>Percentage: %{percent}<extra></extra>'
            )])
            fig_count_pie.update_layout(
                title='Customer Contribution by RFM Category',
                height=500
            )
            st.plotly_chart(fig_count_pie, use_container_width=True)
            
            # 顯示詳細占比
            st.markdown("**Customer占比：**")
            for _, row in category_stats.iterrows():
                st.write(f"- {row['Category']}: {row['Count']:,.0f} ({row['Count_Pct']:.2f}%)")

# 生成退貨分析
def generate_return_analysis(data):
    """生成退貨分析可視化（使用散點圖）"""
    st.markdown("## 🔄 Return Analysis")
    
    if data is None:
        return
    
    # 從MOM數據獲取Return rate和Return amount的月度數據
    mom_df = filter_2011_data(data.get('mom', pd.DataFrame()))
    
    # 顯示Return rate線圖和Return amount柱狀圖（同一張圖，雙Y軸）
    if len(mom_df) > 0 and 'Return_Orders' in mom_df.columns and 'Normal_Orders' in mom_df.columns and 'Return' in mom_df.columns:
        st.markdown("### Return Rate & Return Amount Trends")
        # 計算Return rate
        mom_df['Return_Rate'] = (mom_df['Return_Orders'] / (mom_df['Return_Orders'] + mom_df['Normal_Orders']) * 100).fillna(0)
        # Return amount使用絕對值
        mom_df['Return_Amount'] = mom_df['Return'].abs()
        
        # 創建雙Y軸圖表：Return rate線圖和Return amount柱狀圖
        fig_return_trend = make_subplots(specs=[[{"secondary_y": True}]])
        
        # Return amount柱狀圖（主Y軸）
        fig_return_trend.add_trace(
            go.Bar(
                x=mom_df['YearMonth'],
                y=mom_df['Return_Amount'],
                name='Return Amount',
                marker=dict(color='#e74c3c'),
                opacity=0.7
            ),
            row=1, col=1, secondary_y=False
        )
        
        # Return rate線圖（次Y軸）
        fig_return_trend.add_trace(
            go.Scatter(
                x=mom_df['YearMonth'],
                y=mom_df['Return_Rate'],
                name='Return Rate',
                line=dict(color='#3498db', width=3),
                mode='lines+markers'
            ),
            row=1, col=1, secondary_y=True
        )
        
        fig_return_trend.update_xaxes(title_text="Month", row=1, col=1)
        fig_return_trend.update_yaxes(title_text="Return Amount ($)", row=1, col=1, secondary_y=False)
        fig_return_trend.update_yaxes(title_text="Return Rate (%)", row=1, col=1, secondary_y=True)
        fig_return_trend.update_layout(
            title="Return Rate & Return Amount Trends",
            height=500,
            showlegend=True,
            hovermode='x unified'
        )
        st.plotly_chart(fig_return_trend, use_container_width=True)
    else:
        st.info("ℹ️ 無法顯示Return Rate和Return Amount趨勢：缺少MOM數據或必要列")
    
    # 顯示產品和客戶退貨分析散點圖
    st.markdown("### Product & Customer Return Analysis")
    
    return_product_df = data.get('return_product', pd.DataFrame())
    
    # 讀取客戶退貨數據（如果存在）
    try:
        return_abnormal_file = 'Return and Abnormal_2011_11.xlsx'
        return_customer_df = pd.read_excel(return_abnormal_file, sheet_name='Return analysis customer')
    except:
        try:
            return_abnormal_file = 'Return and Abnormal.xlsx'
            return_customer_df = pd.read_excel(return_abnormal_file, sheet_name='Return analysis customer')
        except:
            return_customer_df = pd.DataFrame()
    
    col1, col2 = st.columns(2)
    
    # 左邊：產品退貨分析散點圖
    with col1:
        if len(return_product_df) > 0 and 'Return_Amount' in return_product_df.columns and 'Return_Rate' in return_product_df.columns:
            # 確保有 Category 列，如果沒有則創建
            if 'Category' not in return_product_df.columns:
                return_product_df['Category'] = 'Unknown'
            
            fig_product = px.scatter(
                return_product_df,
                x='Return_Amount',
                y='Return_Rate',
                color='Category',
                size='Return_Count',
                hover_data=['StockCode', 'Return_Count'],
                title='Product Return Analysis',
                labels={
                    'Return_Amount': 'Return Amount',
                    'Return_Rate': 'Return Rate',
                    'Category': 'Category'
                },
                color_discrete_map={
                    'High-return items': '#e74c3c',
                    'Medium-return items': '#f39c12',
                    'Low-return items': '#2ecc71',
                    '100% return items(outlier)': '#8e44ad',
                    'Unknown': '#95a5a6'
                }
            )
            fig_product.update_layout(height=500)
            st.plotly_chart(fig_product, use_container_width=True)
        else:
            st.warning("沒有產品退貨數據或缺少必要列")
    
    # 右邊：客戶退貨分析散點圖
    with col2:
        if len(return_customer_df) > 0 and 'Return_Amount' in return_customer_df.columns and 'Return_Rate' in return_customer_df.columns:
            # 確保有 Category 列，如果沒有則創建
            if 'Category' not in return_customer_df.columns:
                return_customer_df['Category'] = 'Unknown'
            
            # 獲取 CustomerID 列名
            customer_id_col = None
            for col in return_customer_df.columns:
                if 'customer' in col.lower() or 'id' in col.lower():
                    customer_id_col = col
                    break
            
            hover_data = [customer_id_col, 'Return_Count'] if customer_id_col else ['Return_Count']
            
            fig_customer = px.scatter(
                return_customer_df,
                x='Return_Amount',
                y='Return_Rate',
                color='Category',
                size='Return_Count',
                hover_data=hover_data,
                title='Customer Return Analysis',
                labels={
                    'Return_Amount': 'Return Amount',
                    'Return_Rate': 'Return Rate',
                    'Category': 'Category'
                },
                color_discrete_map={
                    'High-return customer': '#e74c3c',
                    'Medium-return customer': '#f39c12',
                    'Low-return customer': '#2ecc71',
                    '100% return customer(outlier)': '#8e44ad',
                    'Unknown': '#95a5a6'
                }
            )
            fig_customer.update_layout(height=500)
            st.plotly_chart(fig_customer, use_container_width=True)
        else:
            st.info("沒有客戶退貨數據（可選）")

# 生成可執行洞察
def generate_insights(data):
    """自動生成可執行洞察"""
    st.markdown("## 💡 Actionable Insights - 2011/11")
    
    insights = []
    
    if data is None:
        return
    
    mom_df = filter_2011_data(data['mom'])
    rfm_df = data['rfm']
    return_product_df = data['return_product']
    
    # 洞察1：異常退貨高峰月份
    if len(mom_df) > 0 and 'Return_Orders' in mom_df.columns:
        # 計算每個月的退貨率
        if 'Normal_Orders' in mom_df.columns:
            mom_df['Monthly_Return_Rate'] = (mom_df['Return_Orders'] / 
                                             (mom_df['Return_Orders'] + mom_df['Normal_Orders']) * 100)
            avg_return_rate = mom_df['Monthly_Return_Rate'].mean()
            high_return_months = mom_df[mom_df['Monthly_Return_Rate'] > avg_return_rate * 1.5]
            
            if len(high_return_months) > 0:
                months_str = ', '.join(high_return_months['YearMonth'].astype(str).tolist())
                insights.append(f"⚠️ **異常退貨高峰月份**: {months_str} 的退貨率明顯高於平均水平")
    
    # 洞察2：客戶活動下降的細分
    if len(rfm_df) > 0 and 'Category' in rfm_df.columns:
        at_risk_count = len(rfm_df[rfm_df['Category'] == 'At Risk'])
        lost_count = len(rfm_df[rfm_df['Category'] == 'Lost'])
        total_customers = len(rfm_df)
        
        if at_risk_count + lost_count > total_customers * 0.3:
            insights.append(f"📉 **客戶流失風險**: {at_risk_count + lost_count} 個客戶（{((at_risk_count + lost_count)/total_customers*100):.1f}%）處於'At Risk'或'Lost'狀態，需要立即採取保留措施")
    
    # 洞察3：造成最多收入損失的產品
    if len(return_product_df) > 0 and 'Return_Amount' in return_product_df.columns:
        top_loss_products = return_product_df.nlargest(5, 'Return_Amount')
        if len(top_loss_products) > 0:
            products_str = ', '.join(top_loss_products['StockCode'].astype(str).tolist())
            total_loss = top_loss_products['Return_Amount'].sum()
            insights.append(f"💰 **高損失產品**: 產品 {products_str} 造成了最多的退貨損失（總計 ${abs(total_loss):,.0f}），建議檢查產品質量或客戶服務流程")
    
    # 顯示洞察
    if len(insights) > 0:
        for i, insight in enumerate(insights, 1):
            st.markdown(f"""
            <div class="insight-box">
                <strong>洞察 {i}:</strong><br>
                {insight}
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("暫時沒有可用的洞察")

# 主函數
def main():
    # 顯示標題
    st.title("📊 E-commerce Dashboard")
    st.markdown("**數據分析時間範圍: 2011年1月 - 2011年11月**")
    
    # 顯示加載狀態
    with st.spinner("正在加載數據..."):
        data = load_data()
    
    if data is None:
        st.error("❌ 無法加載數據")
        st.markdown("### 請確保以下文件存在：")
        st.write("1. **彙總表.xlsx** - 必須包含以下工作表：")
        st.write("   - MOM")
        st.write("   - AOV_ARPU")
        st.write("   - RFM")
        st.write("   - SKU")
        st.write("   - Sales by Country")
        st.write("2. **Return and Abnormal_2011_11.xlsx** (可選)")
        st.write("   - Return analysis product")
        st.write("   - Abnormal analysis product")
        st.markdown("---")
        st.info("💡 提示: 請先運行 `execute_prompt.py` 生成 彙總表.xlsx")
        return
    
    # 檢查關鍵數據
    if len(data.get('mom', pd.DataFrame())) == 0:
        st.warning("⚠️ MOM 數據為空，無法顯示大部分圖表")
        st.info("請檢查 彙總表.xlsx 是否包含 MOM 工作表")
        return
    
    # 生成KPI
    try:
        generate_kpi(data)
    except Exception as e:
        st.error(f"生成 KPI 時發生錯誤: {e}")
        import traceback
        st.code(traceback.format_exc())
    
    st.divider()
    
    # 生成月度趨勢
    try:
        generate_mom_charts(data)
    except Exception as e:
        st.error(f"生成月度趨勢圖表時發生錯誤: {e}")
        import traceback
        st.code(traceback.format_exc())
    
    st.divider()
    
    # 生成RFM可視化
    try:
        generate_rfm_visualization(data)
    except Exception as e:
        st.error(f"生成 RFM 可視化時發生錯誤: {e}")
        import traceback
        st.code(traceback.format_exc())
    
    st.divider()
    
    # 生成退貨分析
    try:
        generate_return_analysis(data)
    except Exception as e:
        st.error(f"生成退貨分析時發生錯誤: {e}")
        import traceback
        st.code(traceback.format_exc())
    
    st.divider()
    
    # 生成洞察
    try:
        generate_insights(data)
    except Exception as e:
        st.error(f"生成洞察時發生錯誤: {e}")
        import traceback
        st.code(traceback.format_exc())

if __name__ == "__main__":
    main()

