import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from collections import OrderedDict
import random
import io
import re
import sys

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
if 'analysis_completed' not in st.session_state:
    st.session_state.analysis_completed = False
if 'select_all_trigger' not in st.session_state:
    st.session_state.select_all_trigger = False
if 'multiselect_key' not in st.session_state:
    st.session_state.multiselect_key = 0

st.title("🏭 Localhost MRP Dashboard")

tab1, tab2, tab3 = st.tabs(["📦 Stock & PO Management", "🧪 FG Formulas & Settings", "📊 Production Planning"])

# Function to generate or get color for FG code
def get_fg_color(fg_code):
    if fg_code not in st.session_state.fg_colors:
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
            arrival_date = row['Arrival Date']
            arrival_str = arrival_date.strftime('%d/%m/%Y') if hasattr(arrival_date, 'strftime') else str(arrival_date)
            
            html_content += f"""
                <tr>
                    <td>{row['RM Code']}</td>
                    <td>{row['Quantity']:,.4f} Kg</td>
                    <td>{arrival_str}</td>
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
                arrival_date = row['Arrival Date']
                arrival_str = arrival_date.strftime('%d/%m/%Y') if hasattr(arrival_date, 'strftime') else str(arrival_date)
                    
                po_data.append([
                    str(row['RM Code']),
                    f"{row['Quantity']:,.4f} Kg",
                    arrival_str,
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
    except Exception as e:
        st.error(f"PDF generation error: {str(e)}")
        html_content = generate_html_report(results, shortage_details, prod_date, total_volume, ready_fgs, delayed_pos, po_status)
        return html_content.encode('utf-8'), "html"

# NEW: Improved function to generate missing RM data based on Expected Capacities
def generate_missing_rm_summary_from_results(results, shortage_details, fg_expected_capacity, fg_formulas, calculation_margin):
    """Generate missing RM summary based on actual production results and expected capacities"""
    missing_data = []
    
    # First, create a dictionary of actual production for each FG
    fg_production = {}
    for item in results:
        fg_code = item['FG']
        # Extract actual capacity from the string "X,XXX.X Kg"
        actual_str = item['Actual'].replace(' Kg', '').replace(',', '')
        try:
            actual_capacity = float(actual_str)
            fg_production[fg_code] = actual_capacity
        except:
            fg_production[fg_code] = 0
    
    # Process each FG that has shortage details
    for fg_code, shortage_items in shortage_details.items():
        if not shortage_items:
            continue
            
        # Get expected capacity for this FG
        expected_capacity = fg_expected_capacity.get(fg_code, 0)
        actual_capacity = fg_production.get(fg_code, 0)
        
        # Calculate batches based on actual production (not expected)
        actual_batches = int(actual_capacity // 25) if actual_capacity >= 25 else 0
        
        # Get formula for this FG
        formula = fg_formulas[fg_formulas['FG Code'] == fg_code]
        if formula.empty:
            continue
        
        for item in shortage_items:
            if "Shortage" in item:
                try:
                    # Extract RM Code
                    rm_code = item.split(":")[0].strip()
                    
                    # Find this RM in the formula
                    rm_row = formula[formula['RM Code'] == rm_code]
                    if rm_row.empty:
                        continue
                    
                    req_per_batch = float(rm_row['Quantity'].iloc[0])
                    
                    # Parse the shortage string
                    # Format: "RM123: Required X.XXXX Kg, Available Y.YYYY Kg, Shortage Z.ZZZZ Kg"
                    shortage_match = re.search(r'Shortage ([\d.]+)', item)
                    if shortage_match:
                        shortage_qty = float(shortage_match.group(1))
                        
                        # Calculate total required based on what we tried to produce
                        # If expected capacity > 0, use that to calculate required
                        if expected_capacity > 0:
                            expected_batches = max(1, int(expected_capacity // 25))
                            total_required = req_per_batch * expected_batches
                        else:
                            # If no expected capacity, use maximum possible
                            total_required = req_per_batch * actual_batches if actual_batches > 0 else 0
                        
                        # Get available from stock (required - shortage)
                        available_qty = total_required - shortage_qty if total_required > shortage_qty else 0
                        
                        missing_data.append({
                            'FG Code': fg_code,
                            'RM Code': rm_code,
                            'Expected Capacity (Kg)': expected_capacity,
                            'Actual Production (Kg)': actual_capacity,
                            'Required per Batch (Kg)': round(req_per_batch, calculation_margin),
                            'Total Required (Kg)': round(total_required, calculation_margin),
                            'Available (Kg)': round(available_qty, calculation_margin),
                            'Shortage (Kg)': round(shortage_qty, calculation_margin)
                        })
                except Exception as e:
                    # If parsing fails, skip this item
                    continue
    
    if missing_data:
        # Create detailed DataFrame
        detailed_df = pd.DataFrame(missing_data)
        
        # Create summary DataFrame (group by RM Code)
        if not detailed_df.empty:
            summary_data = []
            for rm_code in detailed_df['RM Code'].unique():
                rm_rows = detailed_df[detailed_df['RM Code'] == rm_code]
                total_shortage = rm_rows['Shortage (Kg)'].sum()
                total_required = rm_rows['Total Required (Kg)'].sum()
                total_available = rm_rows['Available (Kg)'].sum()
                affected_fgs = ', '.join(rm_rows['FG Code'].unique())
                
                summary_data.append({
                    'RM Code': rm_code,
                    'Total Required (Kg)': round(total_required, calculation_margin),
                    'Total Available (Kg)': round(total_available, calculation_margin),
                    'Total Shortage (Kg)': round(total_shortage, calculation_margin),
                    'Affected FG Codes': affected_fgs,
                    'Number of Affected FGs': len(rm_rows['FG Code'].unique())
                })
            
            summary_df = pd.DataFrame(summary_data)
            return detailed_df, summary_df
    
    return pd.DataFrame(), pd.DataFrame()

# Function to generate Excel file with multiple sheets
def generate_missing_rm_excel(detailed_df, summary_df, shortage_details, results, prod_date):
    """Generate Excel file with multiple sheets for missing RM analysis"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Detailed Missing RM by FG
        if not detailed_df.empty:
            detailed_df.to_excel(writer, sheet_name='Missing RM Detailed', index=False)
        
        # Sheet 2: Summary by RM Code
        if not summary_df.empty:
            summary_df.to_excel(writer, sheet_name='Missing RM Summary', index=False)
        
        # Sheet 3: FG Production Status
        if results:
            status_df = pd.DataFrame(results)
            status_df.to_excel(writer, sheet_name='FG Production Status', index=False)
        
        # Sheet 4: Raw Shortage Details
        shortage_list = []
        for fg_code, items in shortage_details.items():
            for item in items:
                shortage_list.append({
                    'FG Code': fg_code,
                    'Shortage Details': item
                })
        
        if shortage_list:
            shortage_raw_df = pd.DataFrame(shortage_list)
            shortage_raw_df.to_excel(writer, sheet_name='Raw Shortage Data', index=False)
    
    output.seek(0)
    return output

# Function to generate All Missing RM Report
def generate_all_missing_rm_report(detailed_df, summary_df, prod_date):
    """Generate a comprehensive Excel report with all missing RM data"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Executive Summary
        exec_summary = pd.DataFrame({
            'Report Type': ['Missing RM Analysis Report'],
            'Production Date': [prod_date.strftime('%Y-%m-%d')],
            'Report Generated': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
            'Total Missing RM Types': [len(summary_df) if not summary_df.empty else 0],
            'Total Shortage Quantity (Kg)': [summary_df['Total Shortage (Kg)'].sum() if not summary_df.empty else 0],
            'Total Affected FGs': [len(set(detailed_df['FG Code'])) if not detailed_df.empty else 0]
        })
        exec_summary.to_excel(writer, sheet_name='Executive Summary', index=False)
        
        # Sheet 2: RM Summary (Priority View)
        if not summary_df.empty:
            priority_summary = summary_df.copy()
            priority_summary = priority_summary.sort_values('Total Shortage (Kg)', ascending=False)
            priority_summary['Priority'] = range(1, len(priority_summary) + 1)
            priority_summary = priority_summary[['Priority', 'RM Code', 'Total Shortage (Kg)', 
                                                'Total Required (Kg)', 'Total Available (Kg)', 
                                                'Number of Affected FGs', 'Affected FG Codes']]
            priority_summary.to_excel(writer, sheet_name='RM Priority List', index=False)
        
        # Sheet 3: Detailed FG Breakdown
        if not detailed_df.empty:
            detailed_sorted = detailed_df.sort_values(['RM Code', 'FG Code'])
            detailed_sorted.to_excel(writer, sheet_name='Detailed Analysis', index=False)
        
        # Sheet 4: Action Required
        if not summary_df.empty:
            action_items = []
            for _, row in summary_df.iterrows():
                action_items.append({
                    'RM Code': row['RM Code'],
                    'Shortage (Kg)': row['Total Shortage (Kg)'],
                    'Action Required': f"Procure {row['Total Shortage (Kg)']:,.4f} Kg of {row['RM Code']}",
                    'Priority': 'High' if row['Total Shortage (Kg)'] > 100 else 'Medium' if row['Total Shortage (Kg)'] > 50 else 'Low',
                    'Affected Production': f"{row['Number of Affected FGs']} FG(s): {row['Affected FG Codes']}"
                })
            action_df = pd.DataFrame(action_items)
            action_df.to_excel(writer, sheet_name='Action Items', index=False)
    
    output.seek(0)
    return output

# NEW: Function to generate shortage details table for Excel export
def generate_shortage_details_table(shortage_details, calculation_margin):
    """Generate a table with FG code, RM code, Required, and Available columns from shortage details"""
    shortage_data = []
    
    for fg_code, items in shortage_details.items():
        for item in items:
            if "Shortage" in item or "Required" in item:
                try:
                    # Parse the item string
                    # Format examples:
                    # "RM123: Required 100.0000 Kg for 4 batches, Available 50.0000 Kg, Shortage 50.0000 Kg"
                    # "RM123: Required 25.0000 Kg per batch, Available 0.0000 Kg, Shortage 25.0000 Kg"
                    
                    # Extract RM Code
                    rm_code = item.split(":")[0].strip()
                    
                    # Extract Required quantity
                    required_match = re.search(r'Required ([\d.]+)', item)
                    required_qty = float(required_match.group(1)) if required_match else 0
                    
                    # Extract Available quantity
                    available_match = re.search(r'Available ([\d.]+)', item)
                    available_qty = float(available_match.group(1)) if available_match else 0
                    
                    shortage_data.append({
                        'FG Code': fg_code,
                        'RM Code': rm_code,
                        'Required (Kg)': round(required_qty, calculation_margin),
                        'Available (Kg)': round(available_qty, calculation_margin),
                        'Shortage (Kg)': round(required_qty - available_qty, calculation_margin)
                    })
                except Exception as e:
                    # If parsing fails, skip this item
                    continue
    
    if shortage_data:
        return pd.DataFrame(shortage_data)
    else:
        return pd.DataFrame(columns=['FG Code', 'RM Code', 'Required (Kg)', 'Available (Kg)', 'Shortage (Kg)'])

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
            st.session_state.analysis_completed = False
            st.success("RM Stock cleared!")
        
        rm_file = st.file_uploader("Upload RM Stock Excel", type=['xlsx', 'xls'], key="rm_up")
        
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
                        st.success(f"✅ Successfully loaded {len(processed_df)} RM stock records!")
                    else:
                        st.warning("No valid data found in the uploaded file")
                        
                else:
                    missing_cols = []
                    if 'RM Code' not in column_mapping:
                        missing_cols.append("RM Code")
                    if 'Quantity' not in column_mapping:
                        missing_cols.append("Quantity")
                    st.error(f"Missing columns: {', '.join(missing_cols)}. Found columns: {list(df.columns)}")
                    
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
            st.info("📤 No RM stock data loaded yet. Please upload an Excel file with RM Code and Quantity columns.")

    with col2:
        st.subheader("RM in Purchase Orders (PO)")
        
        if st.button("🔄 Clear RM PO", key="clear_po"):
            st.session_state.rm_po = pd.DataFrame(columns=['RM Code', 'Quantity', 'Arrival Date'])
            st.session_state.analysis_completed = False
            st.success("RM PO cleared!")
        
        po_file = st.file_uploader("Upload RM PO Excel", type=['xlsx', 'xls'], key="po_up")
        
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
                        st.success(f"✅ Successfully loaded {len(processed_df)} PO records!")
                    else:
                        st.warning("No valid data found in the uploaded file")
                        
                else:
                    missing_cols = []
                    for col in ['RM Code', 'Quantity', 'Arrival Date']:
                        if col not in column_mapping:
                            missing_cols.append(col)
                    st.error(f"Missing columns: {', '.join(missing_cols)}. Found columns: {list(df_po.columns)}")
                    
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
            st.info("📤 No PO data loaded yet. Please upload an Excel file with RM Code, Quantity, and Arrival Date columns.")
    
    add_footer()

# --- PAGE 2: FG FORMULAS & SETTINGS ---
with tab2:
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("Finished Goods Formula Management")
        
        fg_files = st.file_uploader(
            "Upload FG Formulas (Multiple allowed)", 
            type=['xlsx', 'xls'], 
            accept_multiple_files=True,
            key="fg_uploader"
        )
        
        if fg_files:
            total_loaded = 0
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
                            
                            total_loaded += len(processed_fg['FG Code'].unique())
                            st.success(f"✅ Loaded {len(processed_fg['FG Code'].unique())} FG formulas from {f.name}")
                        else:
                            st.warning(f"No valid data found in {f.name}")
                            
                    else:
                        missing_cols = []
                        for col in ['FG Code', 'RM Code', 'Quantity']:
                            if col not in column_mapping:
                                missing_cols.append(col)
                        st.error(f"{f.name}: Missing columns {', '.join(missing_cols)}. Found: {list(new_fg.columns)}")
                        
                except Exception as e:
                    st.error(f"Error reading {f.name}: {str(e)}")
            
            if total_loaded > 0:
                st.session_state.analysis_completed = False
        
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
            
            # FIXED SELECT ALL FUNCTIONALITY
            # Create two columns for layout
            col_select, col_button = st.columns([4, 1])
            
            with col_select:
                # Get current selection
                if st.session_state.select_all_trigger:
                    current_selection = fg_codes.copy()
                else:
                    current_selection = list(st.session_state.fg_analysis_order.keys())
                
                # Create the multiselect widget with dynamic key
                selected_fgs = st.multiselect(
                    "Select FG Codes for analysis (Order determines FIFO priority):",
                    options=fg_codes,
                    default=current_selection,
                    key=f"fg_analysis_select_{st.session_state.multiselect_key}"
                )
            
            with col_button:
                st.write("")  # Spacing
                st.write("")  # Spacing
                if st.button("📋 Select All", key="select_all_fg_button", type="secondary"):
                    # Set the trigger and force widget refresh
                    st.session_state.select_all_trigger = True
                    st.session_state.multiselect_key += 1
                    st.rerun()
            
            # Update the analysis order based on current selection
            if selected_fgs:
                # Sort selected FGs alphabetically/numerically
                sorted_selected_fgs = sorted(selected_fgs)
                
                # Create new order preserving sorted order
                new_order = OrderedDict()
                for i, fg in enumerate(sorted_selected_fgs):
                    new_order[fg] = i
                
                # Update session state
                st.session_state.fg_analysis_order = new_order
                st.session_state.analysis_completed = False
                
                # Reset select_all_trigger if not all are selected
                if set(selected_fgs) != set(fg_codes):
                    st.session_state.select_all_trigger = False
            else:
                # Clear if nothing selected
                st.session_state.fg_analysis_order = OrderedDict()
                st.session_state.analysis_completed = False
                st.session_state.select_all_trigger = False
            
            # Display the current FIFO order
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
                if fg_codes:  # Only show if there are FGs available
                    st.info("Select FG codes above to analyze production planning in Tab 3")
        
        else:
            st.info("🧪 No FG formulas loaded yet. Please upload Excel files with FG Code, RM Code, and Quantity columns.")
    
    with col2:
        st.subheader("⚙️ Calculation Settings")
        
        st.write("### 🔢 Decimal Precision")
        margin = st.number_input(
            "Decimal places for calculations:",
            min_value=0,
            max_value=6,
            value=st.session_state.calculation_margin,
            help="Controls rounding precision for all calculations",
            key="decimal_precision"
        )
        
        if margin != st.session_state.calculation_margin:
            st.session_state.calculation_margin = int(margin)
        
        st.divider()
        
        st.write("### 🗑️ Data Management")
        
        if st.button("🗑️ Clear All FG Formulas", type="secondary", key="clear_all_fg"):
            st.session_state.fg_formulas = pd.DataFrame(columns=['FG Code', 'RM Code', 'Quantity'])
            st.session_state.fg_analysis_order = OrderedDict()
            st.session_state.fg_expected_capacity = {}
            st.session_state.fg_colors = {}
            st.session_state.analysis_completed = False
            st.session_state.select_all_trigger = False
            st.session_state.multiselect_key = 0
            st.success("All FG formulas cleared!")
        
        st.divider()
        
        if not st.session_state.fg_formulas.empty:
            st.write("### 🎯 Delete Specific FG")
            fg_codes = st.session_state.fg_formulas['FG Code'].unique().tolist()
            to_delete = st.multiselect("Select FG to delete:", fg_codes, key="fg_delete_select")
            
            if st.button("🗑️ Delete Selected FG", type="primary", key="delete_fg") and to_delete:
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
                
                # Update analysis flag
                if not st.session_state.fg_analysis_order:
                    st.session_state.analysis_completed = False
                
                # Reset select all trigger
                st.session_state.select_all_trigger = False
                st.session_state.multiselect_key += 1
                
                st.success(f"✅ Deleted {len(to_delete)} FG(s): {', '.join(to_delete)}")
                st.rerun()
    
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
        if st.button("📊 Generate Production Analysis", type="primary", key="generate_analysis"):
            if st.session_state.fg_analysis_order:
                st.session_state.analysis_completed = True
                st.success("✅ Production analysis generated! Scroll down to see results.")
            else:
                st.warning("⚠️ Please select FG codes in Tab 2 first.")
    
    data_ready = True
    warning_messages = []
    
    if st.session_state.rm_stock.empty:
        warning_messages.append("📦 RM Stock")
        data_ready = False
    
    if st.session_state.fg_formulas.empty:
        warning_messages.append("🧪 FG Formulas")
        data_ready = False
    
    if not st.session_state.fg_analysis_order:
        warning_messages.append("🎯 FG selection for analysis")
        data_ready = False
    
    if not data_ready:
        st.warning(f"⚠️ Please complete the following in previous tabs:")
        for msg in warning_messages:
            st.write(f"- {msg}")
    else:
        # Show analysis if completed
        if st.session_state.analysis_completed:
            decimal_places = st.session_state.calculation_margin
            
            stock_dict = st.session_state.rm_stock.set_index('RM Code')['Quantity'].to_dict()
            stock_dict = {k: round(float(v), decimal_places) for k, v in stock_dict.items()}
            
            # Create a copy for calculation (won't be modified for max capacity calculation)
            initial_stock = stock_dict.copy()
            allocated_stock = stock_dict.copy()
            
            results = []
            shortage_details = {}
            
            for fg in st.session_state.fg_analysis_order.keys():
                formula = st.session_state.fg_formulas[st.session_state.fg_formulas['FG Code'] == fg]
                
                if formula.empty:
                    continue
                
                expected_capacity = st.session_state.fg_expected_capacity.get(fg, 0)
                
                # Calculate MAX capacity first (using initial stock, not allocated stock)
                max_possible_batches_list = []
                max_shortage_breakdown = []
                
                for _, row in formula.iterrows():
                    rm = str(row['RM Code']).strip()
                    req_per_batch = round(float(row['Quantity']), decimal_places)
                    avail = initial_stock.get(rm, 0)  # Use initial stock for max calculation
                    
                    if req_per_batch <= 0 or avail <= 0:
                        max_possible_batches_list.append(0)
                        if avail <= 0:
                            max_shortage_breakdown.append(f"{rm}: Available 0.0000 Kg")
                    else:
                        max_batches_for_rm = int(avail // req_per_batch)
                        max_possible_batches_list.append(max_batches_for_rm)
                
                max_possible_batches = min(max_possible_batches_list) if max_possible_batches_list else 0
                max_capacity = max_possible_batches * 25
                
                # Now calculate ACTUAL capacity based on expected and allocated stock
                possible_batches = []
                missing_rms = []
                shortage_breakdown = []
                
                # Calculate based on expected capacity
                if expected_capacity > 0:
                    # Calculate required batches based on expected capacity
                    expected_batches = max(1, int(expected_capacity // 25))
                    actual_expected_capacity = expected_batches * 25
                    
                    for _, row in formula.iterrows():
                        rm = str(row['RM Code']).strip()
                        req_per_batch = round(float(row['Quantity']), decimal_places)
                        avail = allocated_stock.get(rm, 0)
                        
                        # Calculate total required for expected batches
                        total_required = req_per_batch * expected_batches
                        
                        if total_required <= 0:
                            shortage_breakdown.append(f"{rm}: Invalid requirement ({req_per_batch:.{decimal_places}f} Kg per batch)")
                            possible_batches.append(0)
                        elif avail <= 0:
                            possible_batches.append(0)
                            missing_rms.append(rm)
                            shortage_breakdown.append(f"{rm}: Required {total_required:.{decimal_places}f} Kg, Available 0.0000 Kg")
                        else:
                            # Check if we have enough for expected batches
                            if avail >= total_required:
                                max_batches_for_rm = expected_batches
                                possible_batches.append(max_batches_for_rm)
                            else:
                                # Calculate how many batches we can make
                                max_batches_for_rm = int(avail // req_per_batch)
                                possible_batches.append(max_batches_for_rm)
                                
                                if max_batches_for_rm < expected_batches:
                                    missing_rms.append(rm)
                                    shortage = total_required - avail
                                    shortage_breakdown.append(f"{rm}: Required {total_required:.{decimal_places}f} Kg for {expected_batches} batches, Available {avail:.{decimal_places}f} Kg, Shortage {shortage:.{decimal_places}f} Kg")
                else:
                    # If no expected capacity, calculate maximum possible
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
                            shortage_breakdown.append(f"{rm}: Required {req_per_batch:.{decimal_places}f} Kg per batch, Available 0.0000 Kg")
                        else:
                            max_batches_for_rm = int(avail // req_per_batch)
                            possible_batches.append(max_batches_for_rm)
                            
                            if max_batches_for_rm == 0:
                                missing_rms.append(rm)
                                shortage = req_per_batch - avail
                                shortage_breakdown.append(f"{rm}: Required {req_per_batch:.{decimal_places}f} Kg per batch, Available {avail:.{decimal_places}f} Kg, Shortage {shortage:.{decimal_places}f} Kg")
                
                if expected_capacity > 0:
                    # For expected capacity mode, use minimum of possible batches
                    max_possible_for_actual = min(possible_batches) if possible_batches else 0
                    actual_batches = min(expected_batches, max_possible_for_actual)
                    actual_capacity = actual_batches * 25
                else:
                    # For auto mode
                    max_possible_for_actual = min(possible_batches) if possible_batches else 0
                    actual_batches = max_possible_for_actual
                    actual_capacity = max_possible_for_actual * 25
                
                if actual_capacity >= 25:
                    status = "✅ Ready"
                else:
                    status = "❌ Shortage"
                
                shortage_details[fg] = shortage_breakdown
                
                if missing_rms:
                    missing_display = f"{len(missing_rms)} RM(s)"
                else:
                    missing_display = "None"
                
                # Allocate stock for production
                if actual_batches > 0 and status == "✅ Ready":
                    for _, row in formula.iterrows():
                        rm = str(row['RM Code']).strip()
                        req_total = round(row['Quantity'] * actual_batches, decimal_places)
                        if rm in allocated_stock:
                            allocated_stock[rm] = round(allocated_stock[rm] - req_total, decimal_places)
                
                results.append({
                    "FG": fg,
                    "Expected": f"{expected_capacity:,.1f} Kg" if expected_capacity > 0 else "Auto",
                    "Max": f"{max_capacity:,.1f} Kg",  # This is the TRUE max based on initial stock
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
            
            # Generate Missing RM Summary based on Expected Capacities
            detailed_missing_df, summary_missing_df = generate_missing_rm_summary_from_results(
                results, 
                shortage_details,
                st.session_state.fg_expected_capacity,
                st.session_state.fg_formulas,
                st.session_state.calculation_margin
            )
            
            # Generate shortage details table for Excel export
            shortage_table_df = generate_shortage_details_table(shortage_details, st.session_state.calculation_margin)
            
            # Display Missing RM Summary if available
            if summary_missing_df is not None and not summary_missing_df.empty:
                st.divider()
                st.write("### 📋 Missing RM Summary (Based on Expected Capacities)")
                
                # Display summary table
                st.dataframe(
                    summary_missing_df,
                    use_container_width=True,
                    height=min(300, len(summary_missing_df) * 35 + 40),
                    hide_index=True,
                    column_config={
                        "RM Code": st.column_config.TextColumn("RM Code", width="small"),
                        "Total Required (Kg)": st.column_config.NumberColumn("Total Required", format="%.4f"),
                        "Total Available (Kg)": st.column_config.NumberColumn("Total Available", format="%.4f"),
                        "Total Shortage (Kg)": st.column_config.NumberColumn("Total Shortage", format="%.4f"),
                        "Affected FG Codes": st.column_config.TextColumn("Affected FGs", width="medium"),
                        "Number of Affected FGs": st.column_config.NumberColumn("Affected Count")
                    }
                )
            
            st.write("### 🎯 Set Expected Capacities")
            st.caption("Set 0 for automatic calculation based on available RM")
            
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
                st.info("✅ No shortages detected for selected FGs")
            
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
                    chart_data['Max_Num'] = chart_data['Max'].str.replace(' Kg', '').str.replace(',', '').astype(float)
                    
                    color_discrete_map = {}
                    for fg_code in chart_data['FG'].unique():
                        color_discrete_map[fg_code] = get_fg_color(fg_code)
                    
                    # Create stacked bar chart for Actual vs Max
                    fig1_data = []
                    for _, row in chart_data.iterrows():
                        fig1_data.append({'FG': row['FG'], 'Capacity': row['Actual_Num'], 'Type': 'Actual'})
                        fig1_data.append({'FG': row['FG'], 'Capacity': row['Max_Num'] - row['Actual_Num'], 'Type': 'Available'})
                    
                    fig1_df = pd.DataFrame(fig1_data)
                    
                    fig1 = px.bar(
                        fig1_df,
                        x='FG',
                        y='Capacity',
                        title="Actual vs Maximum Capacity",
                        color='Type',
                        color_discrete_map={'Actual': '#2ca02c', 'Available': '#aec7e8'},
                        text=fig1_df['Capacity'].apply(lambda x: f"{x:,.1f}" if x > 0 else ''),
                        hover_data=['Type']
                    )
                    
                    fig1.update_traces(
                        texttemplate='%{text}',
                        textposition='outside',
                        hovertemplate='<b>%{x}</b><br>' +
                                    'Type: %{customdata[0]}<br>' +
                                    'Capacity: %{y:,.1f} Kg<br>' +
                                    '<extra></extra>'
                    )
                    
                    fig1.update_layout(
                        yaxis_title="Capacity (Kg)",
                        showlegend=True,
                        height=400,
                        xaxis_tickangle=-45,
                        legend_title="Capacity Type",
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
                f"**⚙️ Current Settings:**\n"
                f"• Decimal Precision: {st.session_state.calculation_margin} places\n"
                f"• FIFO Order: {', '.join(st.session_state.fg_analysis_order.keys())}\n"
                f"• Batch Size: 25 Kg per batch"
            )
            
            # --- EXPORT REPORTS ---
            st.divider()
            st.write("### 📤 Export Reports")
            
            # Create 3 columns for export buttons
            export_col1, export_col2, export_col3 = st.columns(3)
            
            with export_col1:
                # PDF/HTML Report Button
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
                        file_name = f"MRP_Production_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                        btn_label = "⬇️ PDF Report"
                    else:
                        mime_type = "text/html"
                        file_name = f"MRP_Production_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
                        btn_label = "⬇️ HTML Report"
                    
                    st.download_button(
                        label=btn_label,
                        data=report_data,
                        file_name=file_name,
                        mime=mime_type,
                        key="pdf_html_download"
                    )
                    
                except Exception as e:
                    st.error(f"Error generating report: {str(e)}")
            
            with export_col2:
                # Excel Export Button with Shortage Details only
                if not results:
                    st.info("No production data")
                elif not shortage_table_df.empty:
                    try:
                        # Create Excel with shortage details only
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            # Sheet 1: Shortage Details (FG Code, RM Code, Required, Available)
                            shortage_table_df.to_excel(writer, sheet_name='Shortage Details', index=False)
                            
                            # Sheet 2: Production Summary
                            results_df = pd.DataFrame(results)
                            results_df.to_excel(writer, sheet_name='Production Summary', index=False)
                            
                            # Sheet 3: Missing RM Summary (if available)
                            if not summary_missing_df.empty:
                                summary_missing_df.to_excel(writer, sheet_name='Missing RM Summary', index=False)
                        
                        output.seek(0)
                        
                        st.download_button(
                            label="📈 Shortage Details (Excel)",
                            data=output,
                            file_name=f"Shortage_Details_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="shortage_excel_download"
                        )
                        
                    except Exception as e:
                        st.error(f"Error generating Excel file: {str(e)}")
                else:
                    # Create a simple Excel with production results if no shortages
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        results_df = pd.DataFrame(results)
                        results_df.to_excel(writer, sheet_name='Production Summary', index=False)
                    output.seek(0)
                    
                    st.download_button(
                        label="📈 Production Summary (Excel)",
                        data=output,
                        file_name=f"Production_Summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="results_excel_download"
                    )
            
            with export_col3:
                # Complete Report Button (Missing RM Analysis)
                if not results:
                    st.info("No production data")
                elif not detailed_missing_df.empty and not summary_missing_df.empty:
                    try:
                        all_missing_report = generate_all_missing_rm_report(
                            detailed_missing_df,
                            summary_missing_df,
                            prod_date
                        )
                        
                        st.download_button(
                            label="🚀 Complete Missing RM Report",
                            data=all_missing_report,
                            file_name=f"Complete_Missing_RM_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="all_missing_download"
                        )
                        
                    except Exception as e:
                        st.error(f"Error generating comprehensive report: {str(e)}")
                else:
                    # Create a basic report even without shortages
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        results_df = pd.DataFrame(results)
                        results_df.to_excel(writer, sheet_name='Production Summary', index=False)
                        
                        exec_summary = pd.DataFrame({
                            'Report Type': ['Production Planning Report'],
                            'Production Date': [prod_date.strftime('%Y-%m-%d')],
                            'Report Generated': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                            'Total FG Types': [len(results)],
                            'Total Production Volume (Kg)': [total_volume],
                            'Producible FGs': [len(ready_fgs)]
                        })
                        exec_summary.to_excel(writer, sheet_name='Executive Summary', index=False)
                    output.seek(0)
                    
                    st.download_button(
                        label="📋 Basic Production Report",
                        data=output,
                        file_name=f"Production_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="production_report_download"
                    )
            
            # Display the shortage details table that will be exported
            if not shortage_table_df.empty:
                st.divider()
                st.write("### 📊 Shortage Details for Export")
                
                # Display the table that will be exported
                st.dataframe(
                    shortage_table_df,
                    use_container_width=True,
                    height=min(300, len(shortage_table_df) * 35 + 40),
                    hide_index=True,
                    column_config={
                        "FG Code": st.column_config.TextColumn("FG Code", width="small"),
                        "RM Code": st.column_config.TextColumn("RM Code", width="small"),
                        "Required (Kg)": st.column_config.NumberColumn("Required", format="%.4f"),
                        "Available (Kg)": st.column_config.NumberColumn("Available", format="%.4f"),
                        "Shortage (Kg)": st.column_config.NumberColumn("Shortage", format="%.4f")
                    }
                )
                st.caption("This table will be exported in the 'Shortage Details (Excel)' download")
        else:
            st.info("👆 Click 'Generate Production Analysis' button above to see production planning results.")
    
    add_footer()
