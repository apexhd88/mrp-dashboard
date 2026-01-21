import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from collections import OrderedDict
import random
import io

st.set_page_config(page_title="MRP System Dashboard", layout="wide")

# --- Session State Initialization ---
if 'rm_stock' not in st.session_state:
    st.session_state.rm_stock = pd.DataFrame(columns=['RM Code', 'Quantity'])
if 'rm_po' not in st.session_state:
    st.session_state.rm_po = pd.DataFrame(columns=['RM Code', 'Quantity', 'Arrival Date'])
if 'fg_formulas' not in st.session_state:
    st.session_state.fg_formulas = pd.DataFrame(columns=['FG Code', 'RM Code', 'Quantity'])
if 'fg_analysis_order' not in st.session_state:
    st.session_state.fg_analysis_order = OrderedDict()
if 'fg_expected_capacity' not in st.session_state:
    st.session_state.fg_expected_capacity = {}
if 'calculation_margin' not in st.session_state:
    st.session_state.calculation_margin = 3
if 'fg_colors' not in st.session_state:
    st.session_state.fg_colors = {}
if 'select_all_trigger' not in st.session_state:
    st.session_state.select_all_trigger = False

st.title("🏭 Localhost MRP Dashboard")

tab1, tab2, tab3 = st.tabs(["📦 Stock & PO Management", "🧪 FG Formulas & Settings", "📊 Production Planning"])

# Function to generate or get color for FG code
def get_fg_color(fg_code):
    if fg_code not in st.session_state.fg_colors:
        random.seed(hash(fg_code) % 1000)
        colors_list = [
            '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
            '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
            '#aec7e8', '#ffbb78', '#98df8a', '#ff9896', '#c5b0d5',
            '#c49c94', '#f7b6d2', '#c7c7c7', '#dbdb8d', '#9edae5'
        ]
        color = colors_list[len(st.session_state.fg_colors) % len(colors_list)]
        st.session_state.fg_colors[fg_code] = color
    return st.session_state.fg_colors[fg_code]

# Function to generate HTML report (fallback if PDF fails)
def generate_html_report(results, shortage_details, prod_date, total_volume, ready_fgs, delayed_pos, po_status):
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>MRP Production Planning Summary Report</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 40px; }}
            .header {{ text-align: center; margin-bottom: 30px; }}
            .title {{ font-size: 24px; font-weight: bold; color: #333; }}
            .subtitle {{ font-size: 14px; color: #666; margin-top: 10px; }}
            .section {{ margin: 20px 0; }}
            .section-title {{ font-size: 18px; font-weight: bold; color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 5px; margin-bottom: 15px; }}
            table {{ width: 100%; border-collapse: collapse; margin: 10px 0; }}
            th {{ background-color: #3498db; color: white; padding: 10px; text-align: left; }}
            td {{ padding: 8px; border: 1px solid #ddd; }}
            tr:nth-child(even) {{ background-color: #f2f2f2; }}
            .metric {{ display: inline-block; margin: 10px 20px 10px 0; padding: 10px; background-color: #ecf0f1; border-radius: 5px; }}
            .metric-label {{ font-weight: bold; color: #7f8c8d; }}
            .metric-value {{ font-size: 18px; color: #2c3e50; }}
            .status-ready {{ color: #27ae60; font-weight: bold; }}
            .status-shortage {{ color: #e74c3c; font-weight: bold; }}
            .footer {{ margin-top: 40px; text-align: center; color: #7f8c8d; font-size: 12px; border-top: 1px solid #ddd; padding-top: 20px; }}
            .company-footer {{ margin-top: 30px; text-align: center; color: #3498db; font-weight: bold; font-size: 14px; }}
        </style>
    </head>
    <body>
        <div class="header">
            <div class="title">MRP Production Planning Summary Report</div>
            <div class="subtitle">Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
            <div class="subtitle">Production Date: {prod_date.strftime('%d/%m/%Y')}</div>
        </div>
        
        <div class="section">
            <div class="section-title">Summary Metrics</div>
            <div class="metric">
                <div class="metric-label">Planned Production Date</div>
                <div class="metric-value">{prod_date.strftime('%d/%m/%Y')}</div>
            </div>
            <div class="metric">
                <div class="metric-label">Producible FG Types</div>
                <div class="metric-value">{len(ready_fgs)}</div>
            </div>
            <div class="metric">
                <div class="metric-label">Total Production Volume</div>
                <div class="metric-value">{total_volume:,.1f} Kg</div>
            </div>
            <div class="metric">
                <div class="metric-label">Delayed Purchase Orders</div>
                <div class="metric-value">{delayed_pos}</div>
            </div>
        </div>
    """
    
    # Production Capability List
    if results:
        html_content += """
        <div class="section">
            <div class="section-title">Production Capability List</div>
            <table>
                <tr>
                    <th>FG Code</th>
                    <th>Expected</th>
                    <th>Max (Kg)</th>
                    <th>Actual (Kg)</th>
                    <th>Status</th>
                    <th>Missing RM</th>
                    <th>Batches</th>
                </tr>
        """
        
        for item in results:
            status_class = "status-ready" if "✅" in item['Status'] else "status-shortage"
            html_content += f"""
                <tr>
                    <td>{item['FG']}</td>
                    <td>{item['Expected']}</td>
                    <td>{item['Max']}</td>
                    <td>{item['Actual']}</td>
                    <td class="{status_class}">{item['Status']}</td>
                    <td>{item['Missing']}</td>
                    <td>{item['Batches']}</td>
                </tr>
            """
        
        html_content += """
            </table>
        </div>
        """
    
    # Shortage Details
    shortage_exists = False
    for fg in shortage_details:
        if shortage_details[fg]:
            shortage_exists = True
            break
    
    if shortage_exists:
        html_content += """
        <div class="section">
            <div class="section-title">Shortage Details</div>
        """
        
        for fg in shortage_details:
            if shortage_details[fg]:
                html_content += f"""
                <div style="margin: 15px 0;">
                    <div style="font-weight: bold; color: #e74c3c;">FG Code: {fg}</div>
                    <ul style="margin: 5px 0 20px 20px;">
                """
                
                for item in shortage_details[fg]:
                    html_content += f"<li>{item}</li>"
                
                html_content += """
                    </ul>
                </div>
                """
        
        html_content += """
        </div>
        """
    
    # PO Delay Tracker
    if po_status is not None and not po_status.empty:
        html_content += """
        <div class="section">
            <div class="section-title">Purchase Order Delay Status</div>
            <table>
                <tr>
                    <th>RM Code</th>
                    <th>Quantity</th>
                    <th>Arrival Date</th>
                    <th>Status</th>
                </tr>
        """
        
        for _, row in po_status.iterrows():
            status_class = "status-shortage" if row['Status'] == 'Delayed' else ""
            html_content += f"""
                <tr>
                    <td>{row['RM Code']}</td>
                    <td>{row['Quantity']:,.4f} Kg</td>
                    <td>{row['Arrival Date'].strftime('%d/%m/%Y') if hasattr(row['Arrival Date'], 'strftime') else str(row['Arrival Date'])}</td>
                    <td class="{status_class}">{row['Status']}</td>
                </tr>
            """
        
        html_content += """
            </table>
        </div>
        """
    
    # Settings Information
    html_content += f"""
        <div class="section">
            <div class="section-title">System Settings</div>
            <div style="margin: 10px 0;">
                • Decimal Precision: {st.session_state.calculation_margin} places<br>
                • FIFO Order: {', '.join(st.session_state.fg_analysis_order.keys()) if st.session_state.fg_analysis_order else 'Not set'}<br>
                • Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            </div>
        </div>
        
        <div class="company-footer">
            RGI - Supply Chain Department
        </div>
        
        <div class="footer">
            Report generated by MRP Dashboard System<br>
            --- End of Report ---
        </div>
    </body>
    </html>
    """
    
    return html_content

# Function to generate PDF report
def generate_report(results, shortage_details, prod_date, total_volume, ready_fgs, delayed_pos, po_status):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        
        buffer = io.BytesIO()
        
        doc = SimpleDocTemplate(buffer, pagesize=A4, 
                              rightMargin=72, leftMargin=72,
                              topMargin=72, bottomMargin=72)
        
        elements = []
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=18,
            spaceAfter=30,
            alignment=1
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=14,
            spaceAfter=12,
            spaceBefore=12
        )
        
        normal_style = styles['Normal']
        
        elements.append(Paragraph("MRP Production Planning Summary Report", title_style))
        elements.append(Paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", normal_style))
        elements.append(Paragraph(f"Production Date: {prod_date.strftime('%d/%m/%Y')}", normal_style))
        elements.append(Spacer(1, 20))
        
        elements.append(Paragraph("Summary Metrics", heading_style))
        
        summary_data = [
            ["Metric", "Value"],
            ["Planned Production Date", prod_date.strftime('%d/%m/%Y')],
            ["Producible FG Types", str(len(ready_fgs))],
            ["Total Production Volume", f"{total_volume:,.1f} Kg"],
            ["Delayed Purchase Orders", str(delayed_pos)]
        ]
        
        summary_table = Table(summary_data, colWidths=[200, 150])
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))
        elements.append(summary_table)
        elements.append(Spacer(1, 20))
        
        elements.append(Paragraph("Production Capability List", heading_style))
        
        if results:
            table_data = [["FG Code", "Expected", "Max (Kg)", "Actual (Kg)", "Status", "Missing RM", "Batches"]]
            
            for item in results:
                table_data.append([
                    item['FG'],
                    item['Expected'],
                    item['Max'],
                    item['Actual'],
                    item['Status'],
                    item['Missing'],
                    str(item['Batches'])
                ])
            
            prod_table = Table(table_data, colWidths=[70, 60, 60, 60, 60, 70, 50])
            prod_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 9),
            ]))
            elements.append(prod_table)
            elements.append(Spacer(1, 20))
        
        shortage_exists = False
        for fg in shortage_details:
            if shortage_details[fg]:
                shortage_exists = True
                break
        
        if shortage_exists:
            elements.append(Paragraph("Shortage Details", heading_style))
            
            for fg in shortage_details:
                if shortage_details[fg]:
                    elements.append(Paragraph(f"FG Code: {fg}", styles['Heading3']))
                    for item in shortage_details[fg]:
                        elements.append(Paragraph(f"• {item}", normal_style))
                    elements.append(Spacer(1, 10))
            elements.append(Spacer(1, 20))
        
        if po_status is not None and not po_status.empty:
            elements.append(Paragraph("Purchase Order Delay Status", heading_style))
            
            po_data = [["RM Code", "Quantity", "Arrival Date", "Status"]]
            
            for _, row in po_status.iterrows():
                po_data.append([
                    str(row['RM Code']),
                    f"{row['Quantity']:,.4f} Kg",
                    row['Arrival Date'].strftime('%d/%m/%Y') if hasattr(row['Arrival Date'], 'strftime') else str(row['Arrival Date']),
                    row['Status']
                ])
            
            po_table = Table(po_data, colWidths=[80, 80, 80, 80])
            po_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 9),
            ]))
            elements.append(po_table)
            elements.append(Spacer(1, 20))
        
        elements.append(Paragraph("System Settings", heading_style))
        settings_text = f"""
        • Decimal Precision: {st.session_state.calculation_margin} places
        • FIFO Order: {', '.join(st.session_state.fg_analysis_order.keys()) if st.session_state.fg_analysis_order else 'Not set'}
        • Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        """
        elements.append(Paragraph(settings_text, normal_style))
        
        elements.append(Spacer(1, 30))
        elements.append(Paragraph("RGI - Supply Chain Department", ParagraphStyle(
            'CompanyFooter',
            parent=styles['Normal'],
            fontSize=12,
            alignment=1,
            textColor=colors.HexColor('#3498db'),
            spaceBefore=20
        )))
        
        elements.append(Spacer(1, 10))
        elements.append(Paragraph("Report generated by MRP Dashboard System", ParagraphStyle(
            'Footer',
            parent=styles['Normal'],
            fontSize=10,
            alignment=1,
            textColor=colors.grey
        )))
        elements.append(Paragraph("--- End of Report ---", ParagraphStyle(
            'Footer',
            parent=styles['Normal'],
            fontSize=10,
            alignment=1,
            textColor=colors.grey
        )))
        
        doc.build(elements)
        
        pdf_data = buffer.getvalue()
        buffer.close()
        
        return pdf_data, "pdf"
    
    except ImportError:
        html_content = generate_html_report(results, shortage_details, prod_date, total_volume, ready_fgs, delayed_pos, po_status)
        return html_content.encode('utf-8'), "html"

# Function to add footer to all tabs
def add_footer():
    st.markdown("---")
    st.markdown(
        """
        <div style="text-align: center; color: #666; font-size: 12px; padding: 10px;">
            <strong>RGI - Supply Chain Department</strong> | MRP Dashboard System
        </div>
        """,
        unsafe_allow_html=True
    )

# --- PAGE 1: STOCK & PO ---
with tab1:
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Raw Material Stock (RM)")
        
        if st.button("🔄 Clear RM Stock", key="clear_rm"):
            st.session_state.rm_stock = pd.DataFrame(columns=['RM Code', 'Quantity'])
        
        rm_file = st.file_uploader("Upload RM Stock Excel", type=['xlsx'], key="rm_up")
        
        if rm_file is not None:
            try:
                df = pd.read_excel(rm_file)
                df.columns = df.columns.str.strip()
                
                column_mapping = {}
                
                for col in df.columns:
                    col_lower = str(col).lower()
                    if 'rm' in col_lower and ('code' in col_lower or 'id' in col_lower):
                        column_mapping['RM Code'] = col
                        break
                
                for col in df.columns:
                    col_lower = str(col).lower()
                    if 'quantity' in col_lower or 'qty' in col_lower or 'amount' in col_lower:
                        column_mapping['Quantity'] = col
                        break
                
                if 'RM Code' in column_mapping and 'Quantity' in column_mapping:
                    processed_df = pd.DataFrame()
                    processed_df['RM Code'] = df[column_mapping['RM Code']].astype(str).str.strip()
                    processed_df['Quantity'] = pd.to_numeric(df[column_mapping['Quantity']], errors='coerce').fillna(0)
                    
                    processed_df = processed_df[processed_df['RM Code'] != '']
                    processed_df = processed_df[processed_df['RM Code'] != 'nan']
                    processed_df = processed_df.dropna(subset=['RM Code'])
                    
                    if not processed_df.empty:
                        st.session_state.rm_stock = processed_df
                    else:
                        st.warning("No valid data found in the uploaded file")
                        
                else:
                    missing_cols = []
                    if 'RM Code' not in column_mapping:
                        missing_cols.append("RM Code")
                    if 'Quantity' not in column_mapping:
                        missing_cols.append("Quantity")
                    st.error(f"Missing columns: {', '.join(missing_cols)}")
                    
            except Exception as e:
                st.error(f"Error processing RM file: {str(e)}")
        
        if not st.session_state.rm_stock.empty:
            st.write("### 📊 Current Stock Inventory")
            
            display_df = st.session_state.rm_stock.copy()
            display_df['Quantity'] = display_df['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
            
            st.dataframe(
                display_df,
                use_container_width=True,
                height=min(300, len(display_df) * 35 + 40),
                hide_index=True
            )
            
            st.write("### 🔍 View RM Details")
            if not st.session_state.rm_stock.empty:
                sel_rm = st.selectbox(
                    "Select RM Code:",
                    st.session_state.rm_stock['RM Code'].unique(),
                    key="rm_select"
                )
                if sel_rm:
                    qty_row = st.session_state.rm_stock[st.session_state.rm_stock['RM Code'] == sel_rm]
                    if not qty_row.empty:
                        qty = qty_row['Quantity'].values[0]
                        st.metric(label=f"Available {sel_rm}", value=f"{qty:,.4f} Kg")
        else:
            st.info("No RM stock data loaded yet. Please upload an Excel file.")

    with col2:
        st.subheader("RM in Purchase Orders (PO)")
        
        if st.button("🔄 Clear RM PO", key="clear_po"):
            st.session_state.rm_po = pd.DataFrame(columns=['RM Code', 'Quantity', 'Arrival Date'])
        
        po_file = st.file_uploader("Upload RM PO Excel", type=['xlsx'], key="po_up")
        
        if po_file is not None:
            try:
                df_po = pd.read_excel(po_file)
                df_po.columns = df_po.columns.str.strip()
                
                column_mapping = {}
                
                for col in df_po.columns:
                    col_lower = str(col).lower()
                    if 'rm' in col_lower and ('code' in col_lower or 'id' in col_lower):
                        column_mapping['RM Code'] = col
                        break
                
                for col in df_po.columns:
                    col_lower = str(col).lower()
                    if 'quantity' in col_lower or 'qty' in col_lower or 'amount' in col_lower:
                        column_mapping['Quantity'] = col
                        break
                
                for col in df_po.columns:
                    col_lower = str(col).lower()
                    if 'arrival' in col_lower or 'date' in col_lower or 'delivery' in col_lower:
                        column_mapping['Arrival Date'] = col
                        break
                
                if all(col in column_mapping for col in ['RM Code', 'Quantity', 'Arrival Date']):
                    processed_df = pd.DataFrame()
                    processed_df['RM Code'] = df_po[column_mapping['RM Code']].astype(str).str.strip()
                    processed_df['Quantity'] = pd.to_numeric(df_po[column_mapping['Quantity']], errors='coerce').fillna(0)
                    
                    date_col = df_po[column_mapping['Arrival Date']]
                    try:
                        processed_df['Arrival Date'] = pd.to_datetime(date_col, dayfirst=True, errors='coerce')
                    except:
                        try:
                            processed_df['Arrival Date'] = pd.to_datetime(date_col, errors='coerce')
                        except:
                            st.error("Could not parse date column")
                            processed_df['Arrival Date'] = pd.NaT
                    
                    processed_df = processed_df[processed_df['RM Code'] != '']
                    processed_df = processed_df[processed_df['RM Code'] != 'nan']
                    processed_df = processed_df.dropna(subset=['RM Code', 'Arrival Date'])
                    
                    if not processed_df.empty:
                        st.session_state.rm_po = processed_df
                    else:
                        st.warning("No valid data found in the uploaded file")
                        
                else:
                    missing_cols = []
                    for col in ['RM Code', 'Quantity', 'Arrival Date']:
                        if col not in column_mapping:
                            missing_cols.append(col)
                    st.error(f"Missing columns: {', '.join(missing_cols)}")
                    
            except Exception as e:
                st.error(f"Error processing PO file: {str(e)}")
        
        if not st.session_state.rm_po.empty:
            st.write("### 📅 PO Schedule")
            
            display_po = st.session_state.rm_po.copy()
            display_po = display_po.sort_values(by='Arrival Date')
            display_po['Quantity'] = display_po['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
            display_po['Arrival Date'] = display_po['Arrival Date'].dt.strftime('%d/%m/%Y')
            
            st.dataframe(
                display_po,
                use_container_width=True,
                height=min(300, len(display_po) * 35 + 40),
                hide_index=True
            )
            
            st.write("### 📊 Total RM in PO by Code")
            if not st.session_state.rm_po.empty:
                total_po = st.session_state.rm_po.groupby('RM Code')['Quantity'].sum().reset_index()
                total_po['Quantity'] = total_po['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
                
                st.dataframe(
                    total_po,
                    use_container_width=True,
                    hide_index=True
                )
        else:
            st.info("No PO data loaded yet. Please upload an Excel file.")
    
    add_footer()

# --- PAGE 2: FG FORMULAS & SETTINGS ---
with tab2:
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("Finished Goods Formula Management")
        
        fg_files = st.file_uploader(
            "Upload FG Formulas (Multiple allowed)", 
            type=['xlsx'], 
            accept_multiple_files=True,
            key="fg_uploader"
        )
        
        if fg_files:
            for f in fg_files:
                try:
                    new_fg = pd.read_excel(f)
                    new_fg.columns = new_fg.columns.str.strip()
                    
                    column_mapping = {}
                    
                    for col in new_fg.columns:
                        col_lower = str(col).lower()
                        if 'fg' in col_lower and ('code' in col_lower or 'id' in col_lower):
                            column_mapping['FG Code'] = col
                            break
                    
                    for col in new_fg.columns:
                        col_lower = str(col).lower()
                        if 'rm' in col_lower and ('code' in col_lower or 'id' in col_lower):
                            column_mapping['RM Code'] = col
                            break
                    
                    for col in new_fg.columns:
                        col_lower = str(col).lower()
                        if 'quantity' in col_lower or 'qty' in col_lower:
                            column_mapping['Quantity'] = col
                            break
                    
                    if all(col in column_mapping for col in ['FG Code', 'RM Code', 'Quantity']):
                        processed_fg = pd.DataFrame()
                        processed_fg['FG Code'] = new_fg[column_mapping['FG Code']].astype(str).str.strip()
                        processed_fg['RM Code'] = new_fg[column_mapping['RM Code']].astype(str).str.strip()
                        processed_fg['Quantity'] = pd.to_numeric(new_fg[column_mapping['Quantity']], errors='coerce').fillna(0)
                        
                        processed_fg = processed_fg[
                            (processed_fg['FG Code'] != '') & 
                            (processed_fg['FG Code'] != 'nan') &
                            (processed_fg['RM Code'] != '') &
                            (processed_fg['RM Code'] != 'nan')
                        ]
                        processed_fg = processed_fg.dropna(subset=['FG Code', 'RM Code'])
                        
                        if not processed_fg.empty:
                            if st.session_state.fg_formulas.empty:
                                st.session_state.fg_formulas = processed_fg
                            else:
                                combined = pd.concat([st.session_state.fg_formulas, processed_fg])
                                st.session_state.fg_formulas = combined.drop_duplicates(
                                    subset=['FG Code', 'RM Code'], 
                                    keep='first'
                                ).reset_index(drop=True)
                            
                            for fg_code in processed_fg['FG Code'].unique():
                                if fg_code not in st.session_state.fg_colors:
                                    get_fg_color(fg_code)
                        else:
                            st.warning(f"No valid data found in {f.name}")
                            
                    else:
                        missing_cols = []
                        for col in ['FG Code', 'RM Code', 'Quantity']:
                            if col not in column_mapping:
                                missing_cols.append(col)
                        st.error(f"{f.name}: Missing columns {', '.join(missing_cols)}")
                        
                except Exception as e:
                    st.error(f"Error reading {f.name}: {str(e)}")
        
        if not st.session_state.fg_formulas.empty:
            st.divider()
            st.write("### 📋 Current FG Formulas")
            
            display_fg = st.session_state.fg_formulas.copy()
            display_fg['Quantity'] = display_fg['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
            
            st.dataframe(
                display_fg,
                use_container_width=True,
                height=min(400, len(display_fg) * 35 + 40),
                hide_index=True
            )
            
            st.divider()
            st.write("### 🔍 Analyze FG Formula")
            
            fg_codes = sorted(st.session_state.fg_formulas['FG Code'].unique().tolist())
            
            if st.session_state.select_all_trigger:
                st.session_state.select_all_trigger = False
                selected_fgs = fg_codes
                st.session_state.fg_analysis_order = OrderedDict()
                for fg in selected_fgs:
                    st.session_state.fg_analysis_order[fg] = len(st.session_state.fg_analysis_order)
                st.rerun()
            else:
                selected_fgs = st.multiselect(
                    "Select FG Codes for analysis (Order determines FIFO priority):",
                    fg_codes,
                    default=list(st.session_state.fg_analysis_order.keys()) if st.session_state.fg_analysis_order else [],
                    key="fg_analysis_select"
                )
            
            col_select, col_button = st.columns([4, 1])
            
            with col_button:
                st.write("")
                st.write("")
                if st.button("📋 Select All", key="select_all_fg"):
                    st.session_state.select_all_trigger = True
                    st.rerun()
            
            if selected_fgs:
                if set(selected_fgs) != set(st.session_state.fg_analysis_order.keys()):
                    st.session_state.fg_analysis_order = OrderedDict()
                
                for fg in selected_fgs:
                    if fg not in st.session_state.fg_analysis_order:
                        st.session_state.fg_analysis_order[fg] = len(st.session_state.fg_analysis_order)
            
            to_remove = [fg for fg in list(st.session_state.fg_analysis_order.keys()) 
                        if fg not in selected_fgs]
            for fg in to_remove:
                del st.session_state.fg_analysis_order[fg]
            
            if st.session_state.fg_analysis_order:
                st.write("**🎯 FIFO Order (First to Last):**")
                for i, (fg, _) in enumerate(st.session_state.fg_analysis_order.items(), 1):
                    st.write(f"{i}. {fg}")
                
                st.write("### 📊 View Formula Details")
                sel_fg_view = st.selectbox(
                    "Select FG to view details:",
                    list(st.session_state.fg_analysis_order.keys()),
                    key="fg_view_select"
                )
                
                if sel_fg_view:
                    formula_view = st.session_state.fg_formulas[
                        st.session_state.fg_formulas['FG Code'] == sel_fg_view
                    ].copy()
                    formula_view['Quantity'] = formula_view['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
                    
                    st.dataframe(
                        formula_view,
                        use_container_width=True,
                        height=min(300, len(formula_view) * 35 + 40),
                        hide_index=True
                    )
        
        else:
            st.info("No FG formulas loaded yet. Please upload Excel files.")
    
    with col2:
        st.subheader("⚙️ Calculation Settings")
        
        st.write("### 🔢 Decimal Precision")
        margin = st.number_input(
            "Decimal places for calculations:",
            min_value=0,
            max_value=6,
            value=st.session_state.calculation_margin,
            help="Controls rounding precision for all calculations"
        )
        
        if margin != st.session_state.calculation_margin:
            st.session_state.calculation_margin = int(margin)
        
        st.divider()
        
        st.write("### 🗑️ Data Management")
        
        if st.button("🗑️ Clear All FG Formulas", type="secondary"):
            st.session_state.fg_formulas = pd.DataFrame(columns=['FG Code', 'RM Code', 'Quantity'])
            st.session_state.fg_analysis_order = OrderedDict()
            st.session_state.fg_expected_capacity = {}
            st.session_state.fg_colors = {}
        
        st.divider()
        
        if not st.session_state.fg_formulas.empty:
            st.write("### 🎯 Delete Specific FG")
            fg_codes = st.session_state.fg_formulas['FG Code'].unique().tolist()
            to_delete = st.multiselect("Select FG to delete:", fg_codes, key="fg_delete_select")
            
            if st.button("🗑️ Delete Selected FG", type="primary") and to_delete:
                st.session_state.fg_formulas = st.session_state.fg_formulas[
                    ~st.session_state.fg_formulas['FG Code'].isin(to_delete)
                ]
                
                for fg in to_delete:
                    if fg in st.session_state.fg_analysis_order:
                        del st.session_state.fg_analysis_order[fg]
                
                for fg in to_delete:
                    if fg in st.session_state.fg_expected_capacity:
                        del st.session_state.fg_expected_capacity[fg]
                
                for fg in to_delete:
                    if fg in st.session_state.fg_colors:
                        del st.session_state.fg_colors[fg]
    
    add_footer()

# --- PAGE 3: PRODUCTION PLANNING ---
with tab3:
    col_date, col_export = st.columns([2, 1])
    
    with col_date:
        st.subheader("Production Planning Summary")
        prod_date = st.date_input("Select Planned Production Date", datetime.now(), key="prod_date")
    
    with col_export:
        st.write("")
        st.write("")
        export_button = st.button("📥 Generate & Download Report", type="primary", key="download_report_top")
    
    data_ready = True
    warning_messages = []
    
    if st.session_state.rm_stock.empty:
        warning_messages.append("RM Stock")
        data_ready = False
    
    if st.session_state.fg_formulas.empty:
        warning_messages.append("FG Formulas")
        data_ready = False
    
    if not st.session_state.fg_analysis_order:
        warning_messages.append("FG selection for analysis")
        data_ready = False
    
    if not data_ready:
        st.warning(f"Please complete the following in previous tabs: {', '.join(warning_messages)}")
    else:
        decimal_places = st.session_state.calculation_margin
        
        stock_dict = st.session_state.rm_stock.set_index('RM Code')['Quantity'].to_dict()
        stock_dict = {k: round(float(v), decimal_places) for k, v in stock_dict.items()}
        allocated_stock = stock_dict.copy()
        
        results = []
        shortage_details = {}
        
        for fg in st.session_state.fg_analysis_order.keys():
            formula = st.session_state.fg_formulas[st.session_state.fg_formulas['FG Code'] == fg]
            
            if formula.empty:
                continue
            
            expected_capacity = st.session_state.fg_expected_capacity.get(fg, 0)
            
            possible_batches = []
            missing_rms = []
            shortage_breakdown = []
            
            for _, row in formula.iterrows():
                rm = str(row['RM Code']).strip()
                req_per_batch = round(float(row['Quantity']), decimal_places)
                avail = allocated_stock.get(rm, 0)
                
                if req_per_batch <= 0:
                    possible_batches.append(0)
                    shortage_breakdown.append(f"{rm}: Invalid requirement ({req_per_batch:.{decimal_places}f} Kg)")
                elif avail <= 0:
                    possible_batches.append(0)
                    missing_rms.append(rm)
                    shortage_breakdown.append(f"{rm}: Required {req_per_batch:.{decimal_places}f} Kg, Available 0.0000 Kg")
                else:
                    max_batches_for_rm = int(avail // req_per_batch)
                    possible_batches.append(max_batches_for_rm)
                    
                    if max_batches_for_rm == 0:
                        missing_rms.append(rm)
                        shortage = req_per_batch - avail
                        shortage_breakdown.append(f"{rm}: Required {req_per_batch:.{decimal_places}f} Kg, Available {avail:.{decimal_places}f} Kg, Shortage {shortage:.{decimal_places}f} Kg")
            
            max_possible_batches = min(possible_batches) if possible_batches else 0
            max_capacity = max_possible_batches * 25
            
            if expected_capacity > 0:
                expected_batches = max(1, int(expected_capacity // 25))
                actual_batches = min(expected_batches, max_possible_batches)
                actual_capacity = actual_batches * 25
            else:
                actual_batches = max_possible_batches
                actual_capacity = max_capacity
            
            if actual_capacity >= 25:
                status = "✅ Ready"
                status_detail = "Sufficient RM"
            else:
                status = "❌ Shortage"
                if max_capacity >= 25:
                    status_detail = f"Insufficient RM for expected {expected_capacity}Kg"
                else:
                    status_detail = f"Insufficient RM for minimum 25Kg batch"
            
            shortage_details[fg] = shortage_breakdown
            
            if missing_rms:
                missing_display = f"{len(missing_rms)} RM(s)"
            else:
                missing_display = "None"
            
            if actual_batches > 0 and status == "✅ Ready":
                for _, row in formula.iterrows():
                    rm = str(row['RM Code']).strip()
                    req_total = round(row['Quantity'] * actual_batches, decimal_places)
                    if rm in allocated_stock:
                        allocated_stock[rm] = round(allocated_stock[rm] - req_total, decimal_places)
            
            results.append({
                "FG": fg,
                "Expected": f"{expected_capacity:,.1f} Kg" if expected_capacity > 0 else "Auto",
                "Max": f"{max_capacity:,.1f} Kg",
                "Actual": f"{actual_capacity:,.1f} Kg",
                "Status": status,
                "Missing": missing_display,
                "Batches": actual_batches
            })
        
        if not st.session_state.rm_po.empty:
            po_status_for_report = st.session_state.rm_po.copy()
            po_status_for_report['Status'] = po_status_for_report['Arrival Date'].apply(
                lambda x: "Delayed" if x.date() < prod_date else "Incoming"
            )
            delayed_pos = len(po_status_for_report[po_status_for_report['Status'] == "Delayed"])
        else:
            po_status_for_report = None
            delayed_pos = 0
        
        ready_fgs = [r for r in results if "✅" in r['Status']]
        total_volume = sum(float(r['Actual'].replace(' Kg', '').replace(',', '')) 
                          for r in results if r['Actual'] != "0.0 Kg")
        
        if export_button:
            try:
                report_data, report_type = generate_report(
                    results, 
                    shortage_details, 
                    prod_date, 
                    total_volume, 
                    ready_fgs, 
                    delayed_pos,
                    po_status_for_report
                )
                
                if report_type == "pdf":
                    mime_type = "application/pdf"
                    file_extension = "pdf"
                    file_name = f"MRP_Production_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                    success_msg = "PDF report generated successfully!"
                else:
                    mime_type = "text/html"
                    file_extension = "html"
                    file_name = f"MRP_Production_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
                    success_msg = "HTML report generated successfully! (PDF library not available)"
                
                st.download_button(
                    label=f"⬇️ Click to Download {file_extension.upper()} Report",
                    data=report_data,
                    file_name=file_name,
                    mime=mime_type,
                    key=f"{file_extension}_download_top"
                )
                
                st.success(success_msg)
                
                if report_type == "html":
                    with st.expander("📄 Report Preview", expanded=False):
                        st.components.v1.html(report_data.decode('utf-8'), height=600, scrolling=True)
                        
            except Exception as e:
                st.error(f"Error generating report: {str(e)}")
        
        st.write("### 🎯 Set Expected Capacities")
        capacity_col1, capacity_col2 = st.columns(2)
        
        fg_list = list(st.session_state.fg_analysis_order.keys())
        mid_point = (len(fg_list) + 1) // 2
        
        with capacity_col1:
            for fg in fg_list[:mid_point]:
                current_val = st.session_state.fg_expected_capacity.get(fg, 0)
                new_capacity = st.number_input(
                    f"{fg} (Kg, min 25):",
                    min_value=0.0,
                    value=float(current_val),
                    step=25.0,
                    format="%.1f",
                    key=f"exp_cap_{fg}"
                )
                if new_capacity != current_val:
                    st.session_state.fg_expected_capacity[fg] = new_capacity
        
        with capacity_col2:
            for fg in fg_list[mid_point:]:
                current_val = st.session_state.fg_expected_capacity.get(fg, 0)
                new_capacity = st.number_input(
                    f"{fg} (Kg, min 25):",
                    min_value=0.0,
                    value=float(current_val),
                    step=25.0,
                    format="%.1f",
                    key=f"exp_cap_{fg}_2"
                )
                if new_capacity != current_val:
                    st.session_state.fg_expected_capacity[fg] = new_capacity
        
        st.divider()
        st.write("### 📊 Production Summary")
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Producible FG", len(ready_fgs))
        col2.metric("Total Volume", f"{total_volume:,.1f} Kg")
        col3.metric("Total Batches", sum(r['Batches'] for r in results))
        col4.metric("Delayed POs", delayed_pos)
        
        st.divider()
        st.write("### 📋 Production Capability List")
        
        res_df = pd.DataFrame(results)
        
        st.dataframe(
            res_df,
            use_container_width=True,
            height=min(400, len(res_df) * 35 + 40),
            hide_index=True,
            column_config={
                "FG": st.column_config.TextColumn("FG Code", width="small"),
                "Expected": st.column_config.TextColumn("Expected", width="small"),
                "Max": st.column_config.TextColumn("Max Cap", width="small"),
                "Actual": st.column_config.TextColumn("Actual Cap", width="small"),
                "Status": st.column_config.TextColumn("Status", width="small"),
                "Missing": st.column_config.TextColumn("Missing RM", width="small"),
                "Batches": st.column_config.NumberColumn("Batches", width="small")
            }
        )
        
        st.divider()
        st.write("### 🔍 Shortage Details")
        
        shortage_exists = False
        for fg in shortage_details:
            if shortage_details[fg]:
                shortage_exists = True
                with st.expander(f"❌ {fg} - RM Shortage Breakdown", expanded=False):
                    for item in shortage_details[fg]:
                        st.write(f"• {item}")
        
        if not shortage_exists:
            st.info("No shortages detected for selected FGs")
        
        if not st.session_state.rm_po.empty:
            st.divider()
            st.write("### ⏰ PO Delay Tracker")
            
            po_display = po_status_for_report.copy()
            po_display['Quantity'] = po_display['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
            po_display['Arrival Date'] = po_display['Arrival Date'].dt.strftime('%d/%m/%Y')
            
            st.dataframe(
                po_display,
                use_container_width=True,
                height=min(300, len(po_display) * 35 + 40),
                hide_index=True,
                column_config={
                    "RM Code": st.column_config.TextColumn("RM Code", width="small"),
                    "Quantity": st.column_config.TextColumn("Quantity", width="small"),
                    "Arrival Date": st.column_config.TextColumn("Arrival", width="small"),
                    "Status": st.column_config.TextColumn("Status", width="small")
                }
            )
        
        st.divider()
        st.write("### 📈 Production Capacity Visualization")
        
        chart_col1, chart_col2 = st.columns(2)
        
        with chart_col1:
            if len(results) > 0:
                res_df = pd.DataFrame(results)
                chart_data = res_df.copy()
                chart_data['Actual_Num'] = chart_data['Actual'].str.replace(' Kg', '').str.replace(',', '').astype(float)
                
                color_discrete_map = {}
                for fg_code in chart_data['FG'].unique():
                    color_discrete_map[fg_code] = get_fg_color(fg_code)
                
                fig1 = px.bar(
                    chart_data,
                    x='FG',
                    y='Actual_Num',
                    title="Actual Production Capacity",
                    color='FG',
                    color_discrete_map=color_discrete_map,
                    text='Actual',
                    hover_data=['Status']
                )
                
                fig1.update_traces(
                    texttemplate='%{text}',
                    textposition='outside',
                    hovertemplate='<b>%{x}</b><br>' +
                                'Capacity: %{y:,.1f} Kg<br>' +
                                'Status: %{customdata[0]}<br>' +
                                '<extra></extra>'
                )
                
                fig1.update_layout(
                    yaxis_title="Capacity (Kg)",
                    showlegend=True,
                    height=400,
                    xaxis_tickangle=-45,
                    legend_title="FG Code",
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    )
                )
                st.plotly_chart(fig1, use_container_width=True)
        
        with chart_col2:
            if len(res_df) > 1:
                pie_data = res_df.copy()
                pie_data['Actual_Num'] = pie_data['Actual'].str.replace(' Kg', '').str.replace(',', '').astype(float)
                pie_data = pie_data[pie_data['Actual_Num'] > 0]
                
                if len(pie_data) > 0:
                    pie_colors = [get_fg_color(fg) for fg in pie_data['FG']]
                    
                    fig2 = px.pie(
                        pie_data,
                        values='Actual_Num',
                        names='FG',
                        title="Capacity Distribution",
                        color='FG',
                        color_discrete_sequence=pie_colors,
                        hole=0.3,
                        hover_data=['Status']
                    )
                    
                    fig2.update_traces(
                        hovertemplate='<b>%{label}</b><br>' +
                                    'Capacity: %{value:,.1f} Kg<br>' +
                                    'Percentage: %{percent}<br>' +
                                    '<extra></extra>'
                    )
                    
                    fig2.update_layout(
                        height=400,
                        legend_title="FG Code",
                        showlegend=True
                    )
                    st.plotly_chart(fig2, use_container_width=True)
        
        st.info(
            f"**Current Settings:**\n"
            f"• Decimal Precision: {st.session_state.calculation_margin} places\n"
            f"• FIFO Order: {', '.join(st.session_state.fg_analysis_order.keys())}"
        )
    
    add_footer()