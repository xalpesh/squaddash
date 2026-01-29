import pandas as pd
import random
import datetime
from dateutil.relativedelta import relativedelta

# ==========================================
# CONFIGURATION
# ==========================================
OUTPUT_FILE = 'Dynamic_Portfolio_Master_v2.xlsx'

# ==========================================
# 1. HELPER: DATE RANGES
# ==========================================
def get_month_columns(start_date, months=12):
    """Generates column headers like 'Jan-26', 'Feb-26'"""
    cols = []
    current = start_date.replace(day=1)
    for _ in range(months):
        cols.append(current)
        current += relativedelta(months=1)
    return cols

# ==========================================
# 2. GENERATE MOCK DATABASE
# ==========================================
def create_enhanced_database():
    # --- 1. Project Master (Added Portfolio & Team) ---
    projects = []
    portfolios = ['Digital Transformation', 'Legacy Modernization', 'Data & AI', 'Cloud Infra']
    teams = ['Squad Alpha', 'Squad Beta', 'Squad Gamma', 'Core Ops']
    
    for i in range(1, 26):
        projects.append({
            'Project_ID': f'P{i:03d}',
            'Project_Name': f'Project {chr(65+(i%26))}-{i:02d}',
            'Portfolio': random.choice(portfolios),
            'Team': random.choice(teams),
            'Lead': random.choice(['Suresh', 'Brandi', 'Abhi', 'Stanley']),
            'PM': random.choice(['Kalpesh', 'Rail', 'Sigmas']),
            'Kickoff_Date': (datetime.date.today() - datetime.timedelta(days=random.randint(20, 100))).strftime('%Y-%m-%d'),
            'End_Date': (datetime.date.today() + datetime.timedelta(days=random.randint(60, 200))).strftime('%Y-%m-%d')
        })
    df_projects = pd.DataFrame(projects)

    # --- 2. Resource Master (With Skills & Manager) ---
    resources = []
    names = ['Mike', 'Sarah', 'Jessica', 'Tom', 'Ravi', 'Siva', 'Wei', 'Lisa', 'John', 'Priya']
    skills = ['Java Fullstack', 'Python/AI', 'SAP ABAP', 'Data Eng', 'DevOps']
    managers = ['Director A', 'Director B', 'Director C']
    
    for idx, name in enumerate(names):
        resources.append({
            'Resource_ID': f'R{idx:03d}',
            'Full_Name': name,
            'Primary_Skill': skills[idx % len(skills)],
            'Team': 'Digital Squad A' if idx % 2 == 0 else 'Core Systems',
            'Manager': random.choice(managers)
        })
    df_resources = pd.DataFrame(resources)

    # --- 3. Allocations (With Start/End Dates) ---
    allocations = []
    for _, proj in df_projects.iterrows():
        # Alloc 2-3 people per project
        team = df_resources.sample(random.randint(2, 3))
        for _, res in team.iterrows():
            start_offset = random.randint(-1, 2)
            duration = random.randint(3, 6)
            s_date = datetime.date.today().replace(day=1) + relativedelta(months=start_offset)
            e_date = s_date + relativedelta(months=duration)
            
            allocations.append({
                'Project_ID': proj['Project_ID'],
                'Resource_ID': res['Resource_ID'],
                'Allocation_%': random.choice([0.5, 1.0]),
                'Start_Date': s_date,
                'End_Date': e_date
            })
    df_allocations = pd.DataFrame(allocations)

    # --- 4. Milestones (Rich Status) ---
    milestones = []
    phases = ['Discovery', 'Design', 'Build', 'UAT', 'Go-Live']
    for p in df_projects['Project_ID']:
        for idx, phase in enumerate(phases):
            date = datetime.date.today() + datetime.timedelta(days=(idx*20) - 20)
            status = 'Completed' if idx < 2 else 'In Progress' if idx == 2 else 'Pending'
            milestones.append({
                'Project_ID': p,
                'Milestone': phase,
                'Date': date.strftime('%Y-%m-%d'),
                'Status': status,
                'Completion_Pct': 1.0 if status == 'Completed' else (0.5 if status == 'In Progress' else 0)
            })
    df_milestones = pd.DataFrame(milestones)

    # --- 5. Weekly Updates (With Upcoming Tasks & Goals) ---
    updates = []
    for p in df_projects['Project_ID']:
        rag = random.choice(['Red', 'Amber', 'Green'])
        updates.append({
            'Project_ID': p,
            'Week': 'Wk 05',
            'RAG': rag,
            'Goal_This_Week': 'Finalize UAT Sign-off and prepare deployment scripts.',
            'Narrative': 'Critical path delayed due to API.' if rag == 'Red' else 'On track.',
            'Upcoming_Tasks': '1. Finalize API Spec\n2. Sign-off UAT Plan\n3. Onboard new QA',
            'Risks': 'None' if rag == 'Green' else 'Resource constraint'
        })
    df_updates = pd.DataFrame(updates)

    # --- 6. SLA ---
    slas = []
    for p in df_projects['Project_ID']:
        slas.append({'Project_ID': p, 'Metric': 'Defect Turnaround', 'Status': random.choice(['Met', 'Met', 'Breached'])})
    df_sla = pd.DataFrame(slas)

    return df_projects, df_resources, df_allocations, df_milestones, df_updates, df_sla

# ==========================================
# 3. AGGREGATE DASHBOARD LOGIC
# ==========================================
def compile_dashboard_data(dfs):
    df_p, df_r, df_a, df_m, df_u, df_s = dfs
    
    dashboard_rows = []

    for _, proj in df_p.iterrows():
        pid = proj['Project_ID']
        
        # 1. Get Weekly Update
        update = df_u[df_u['Project_ID'] == pid].iloc[0]
        
        # 2. Aggregate Milestones
        proj_miles = df_m[df_m['Project_ID'] == pid].sort_values('Date')
        mile_lines = []
        for _, m in proj_miles.iterrows():
            icon = "✅" if m['Status'] == 'Completed' else "⚠️" if m['Status'] == 'In Progress' else "⚪"
            date_str = m['Date'][5:]
            mile_lines.append(f"{icon} {m['Milestone']} ({int(m['Completion_Pct']*100)}%) - {date_str}")
        roadmap_str = "\n".join(mile_lines)
        
        # 3. Aggregate Resources (With Dates)
        proj_alloc = df_a[df_a['Project_ID'] == pid].merge(df_r, on='Resource_ID')
        team_lines = []
        total_util = 0
        for _, r in proj_alloc.iterrows():
            pct = r['Allocation_%']
            s_str = r['Start_Date'].strftime('%b')
            e_str = r['End_Date'].strftime('%b')
            team_lines.append(f"• {r['Full_Name']} ({r['Primary_Skill']}): {int(pct*100)}% [{s_str}-{e_str}]")
            total_util += pct
        team_str = "\n".join(team_lines)
        
        # 4. SLA
        proj_sla = df_s[df_s['Project_ID'] == pid]
        sla_status = "Breached" if "Breached" in proj_sla['Status'].values else "Met"

        # Build Row
        dashboard_rows.append({
            'Project Name': proj['Project_Name'],
            'Portfolio': proj['Portfolio'],
            'Team': proj['Team'],
            'Lead / PM': f"Lead: {proj['Lead']}\nPM: {proj['PM']}",
            'Overall Status': update['RAG'],
            'Milestone Roadmap': roadmap_str,
            'Resource Plan': team_str,
            'Total FTE': total_util,
            'Weekly Update': f"GOAL: {update['Goal_This_Week']}\n\nUPDATE: {update['Narrative']}\n\nTASKS:\n{update['Upcoming_Tasks']}\n\nRISK: {update['Risks']}",
            'SLA': sla_status
        })
        
    # Sort by Portfolio then Team
    df_dashboard = pd.DataFrame(dashboard_rows)
    df_dashboard.sort_values(by=['Portfolio', 'Team'], inplace=True)
    return df_dashboard

# ==========================================
# 4. CALCULATE MONTHLY HEATMAP
# ==========================================
def generate_heatmap_data(df_res, df_alloc):
    today = datetime.date.today().replace(day=1)
    months = get_month_columns(today, 12)
    
    heatmap_rows = []
    
    for _, res in df_res.iterrows():
        rid = res['Resource_ID']
        row = {
            'Resource Name': res['Full_Name'],
            'Primary Skill': res['Primary_Skill'],
            'Manager': res['Manager']
        }
        
        my_allocs = df_alloc[df_alloc['Resource_ID'] == rid]
        
        for m in months:
            month_load = 0
            month_end = m + relativedelta(months=1, days=-1)
            for _, alloc in my_allocs.iterrows():
                if alloc['Start_Date'] <= month_end and alloc['End_Date'] >= m:
                    month_load += alloc['Allocation_%']
            row[m] = month_load
            
        heatmap_rows.append(row)
        
    return pd.DataFrame(heatmap_rows), months

# ==========================================
# 5. EXCEL GENERATION
# ==========================================
def create_workbook():
    dfs = create_enhanced_database()
    df_p, df_r, df_a, df_m, df_u, df_s = dfs
    df_dashboard = compile_dashboard_data(dfs)
    df_heat, month_cols = generate_heatmap_data(df_r, df_a)
    
    writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
    workbook = writer.book

    # --- FORMATS ---
    fmt_header = workbook.add_format({'bold': True, 'fg_color': '#0F2C4C', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    fmt_rich = workbook.add_format({'text_wrap': True, 'valign': 'top', 'border': 1, 'font_size': 10})
    fmt_center = workbook.add_format({'align': 'center', 'valign': 'top', 'border': 1})
    fmt_date = workbook.add_format({'num_format': 'dd-mmm-yy', 'align': 'center', 'border': 1})
    
    # RAG
    rag_red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'align': 'center', 'bold': True, 'border': 1})
    rag_amber = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'align': 'center', 'bold': True, 'border': 1})
    rag_green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'align': 'center', 'bold': True, 'border': 1})
    
    # Heatmap Colors
    fmt_over = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '0%'})
    fmt_free = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'num_format': '0%'})

    # ================= 1. REPORT DASHBOARD =================
    ws_rep = workbook.add_worksheet(">> DASHBOARD <<")
    
    # Write Headers
    headers = list(df_dashboard.columns)
    ws_rep.write_row(0, 0, headers, fmt_header)
    
    # Write Data
    for row_idx, row_data in df_dashboard.iterrows():
        # Using raw index loop since we sorted dataframe
        # Write specific columns with specific formats
        r = list(df_dashboard.itertuples(index=False))
        # Easier loop:
    
    # Let's use pandas to write first, then apply format
    df_dashboard.to_excel(writer, sheet_name=">> DASHBOARD <<", index=False, startrow=1, header=False)
    
    # Apply Column Formats
    ws_rep.set_column('A:A', 25, fmt_rich) # Project
    ws_rep.set_column('B:C', 15, fmt_center) # Portfolio/Team
    ws_rep.set_column('D:D', 15, fmt_rich) # Lead
    ws_rep.set_column('E:E', 10, fmt_center) # Status
    ws_rep.set_column('F:F', 30, fmt_rich) # Roadmap
    ws_rep.set_column('G:G', 30, fmt_rich) # Resource
    ws_rep.set_column('H:H', 8, fmt_center) # FTE
    ws_rep.set_column('I:I', 40, fmt_rich) # Narrative
    ws_rep.set_column('J:J', 10, fmt_center) # SLA
    
    # Conditional Formatting for RAG (Col E - Index 4)
    last_row = len(df_dashboard) + 1
    ws_rep.conditional_format(1, 4, last_row, 4, {'type': 'cell', 'criteria': 'equal to', 'value': '"Red"', 'format': rag_red})
    ws_rep.conditional_format(1, 4, last_row, 4, {'type': 'cell', 'criteria': 'equal to', 'value': '"Amber"', 'format': rag_amber})
    ws_rep.conditional_format(1, 4, last_row, 4, {'type': 'cell', 'criteria': 'equal to', 'value': '"Green"', 'format': rag_green})

    # ================= 2. RESOURCE HEATMAP =================
    df_heat.to_excel(writer, sheet_name=">> RES_HEATMAP <<", index=False)
    ws_heat = writer.sheets[">> RES_HEATMAP <<"]
    ws_heat.write_row(0, 0, df_heat.columns, fmt_header)
    
    # Format Dates in header
    for col_num, val in enumerate(df_heat.columns):
        if isinstance(val, datetime.date):
            ws_heat.write(0, col_num, val.strftime('%b-%y'), fmt_header)
            
    # Conditional Heatmap
    last_row_h = len(df_heat) + 1
    last_col_h = len(df_heat.columns) - 1
    ws_heat.conditional_format(1, 3, last_row_h, last_col_h, {'type': 'cell', 'criteria': '>', 'value': 1, 'format': fmt_over})
    ws_heat.conditional_format(1, 3, last_row_h, last_col_h, {'type': 'cell', 'criteria': '<', 'value': 0.8, 'format': fmt_free})

    # ================= 3. DB TABS =================
    
    # DB_Projects (With Validation)
    df_p.to_excel(writer, sheet_name="DB_Projects", index=False)
    ws_proj = writer.sheets["DB_Projects"]
    # Dropdown for Portfolio (Col C)
    ws_proj.data_validation('C2:C100', {'validate': 'list', 'source': ['Digital Transformation', 'Legacy Modernization', 'Data & AI', 'Cloud Infra']})
    # Dropdown for Team (Col D)
    ws_proj.data_validation('D2:D100', {'validate': 'list', 'source': ['Squad Alpha', 'Squad Beta', 'Squad Gamma', 'Core Ops']})

    # DB_Resources (Validation)
    df_r.to_excel(writer, sheet_name="DB_Resources", index=False)
    ws_res = writer.sheets["DB_Resources"]
    ws_res.data_validation('C2:C100', {'validate': 'list', 'source': ['Java Fullstack', 'Python/AI', 'SAP ABAP', 'Data Eng', 'DevOps']})

    # DB_Allocations
    df_a.to_excel(writer, sheet_name="DB_Allocations", index=False)
    ws_alloc = writer.sheets["DB_Allocations"]
    ws_alloc.set_column('F:G', 12, fmt_date) # Start/End dates

    # Other DBs
    df_u.to_excel(writer, sheet_name="DB_Updates", index=False)
    df_m.to_excel(writer, sheet_name="DB_Milestones", index=False)
    df_s.to_excel(writer, sheet_name="DB_SLA", index=False)
    
    writer.close()
    print(f"✅ Generated Complete v2 Report: {OUTPUT_FILE}")

if __name__ == "__main__":
    create_workbook()