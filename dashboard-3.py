import pandas as pd
import random
import datetime
from dateutil.relativedelta import relativedelta

# ==========================================
# CONFIGURATION
# ==========================================
OUTPUT_FILE = 'Dynamic_Portfolio_Master_v4.xlsx'

# ==========================================
# 1. HELPER FUNCTIONS
# ==========================================
def get_month_columns(start_date, months=12):
    """Generates list of month start dates"""
    cols = []
    current = start_date.replace(day=1)
    for _ in range(months):
        cols.append(current)
        current += relativedelta(months=1)
    return cols

def get_random_date(start_offset_days, duration_days):
    start = datetime.date.today() + datetime.timedelta(days=start_offset_days)
    end = start + datetime.timedelta(days=duration_days)
    return start, end

def get_quarter_list():
    """Generates Q1-2026 to Q4-2033"""
    quarters = []
    for year in range(2026, 2034):
        for q in range(1, 5):
            quarters.append(f"Q{q}-{year}")
    return quarters

# ==========================================
# 2. GENERATE MOCK DATABASE
# ==========================================
def create_database_v4():
    
    # --- 0. Config Data (Quarters) ---
    quarters = get_quarter_list()
    df_config = pd.DataFrame(quarters, columns=['Quarters'])

    # --- 1. DB_Skills (Master List) ---
    skills_data = [
        {'Skill_ID': 'S01', 'Skill_Name': 'Java Fullstack', 'Levels': 'L1, L2, L3'},
        {'Skill_ID': 'S02', 'Skill_Name': 'Python/AI', 'Levels': 'L1, L2, L3'},
        {'Skill_ID': 'S03', 'Skill_Name': 'SAP ABAP', 'Levels': 'L1, L2, L3'},
        {'Skill_ID': 'S04', 'Skill_Name': 'Data Engineering', 'Levels': 'L1, L2, L3'},
        {'Skill_ID': 'S05', 'Skill_Name': 'Cloud/DevOps', 'Levels': 'L1, L2, L3'}
    ]
    df_skills = pd.DataFrame(skills_data)

    # --- 2. DB_Projects (Current Active) ---
    projects = []
    portfolios = ['Digital Transformation', 'Legacy Modernization', 'Data & AI', 'Cloud Infra']
    teams = ['Squad Alpha', 'Squad Beta', 'Squad Gamma', 'Core Ops']
    
    for i in range(1, 26):
        s_date, e_date = get_random_date(-30, 180)
        projects.append({
            'Project_ID': f'P{i:03d}',
            'Project_Name': f'Project {chr(65+(i%26))}-{i:02d}',
            'Portfolio': random.choice(portfolios),
            'Team': random.choice(teams),
            'Goal': random.choice(quarters),
            'Lead': random.choice(['Suresh', 'Brandi', 'Abhi', 'Stanley']),
            'PM': random.choice(['Kalpesh', 'Rail']),
            'Kickoff': s_date,
            'End_Date': e_date
        })
    df_projects = pd.DataFrame(projects)

    # --- 3. DB_Resources (With Skills & Exp) ---
    resources = []
    names = ['Mike', 'Sarah', 'Jessica', 'Tom', 'Ravi', 'Siva', 'Wei', 'Lisa', 'John', 'Priya']
    managers = ['Director A', 'Director B']
    
    for idx, name in enumerate(names):
        skill = random.choice(skills_data)
        resources.append({
            'Resource_ID': f'R{idx:03d}',
            'Full_Name': name,
            'Skill_ID': skill['Skill_ID'],
            'Skill_Level': random.choice(['Junior', 'Standard', 'Senior']),
            'Years_Exp': random.randint(2, 15),
            'Manager': random.choice(managers)
        })
    df_resources = pd.DataFrame(resources)

    # --- 4. DB_Allocations (Active Assignments) ---
    allocations = []
    for _, proj in df_projects.iterrows():
        team_members = df_resources.sample(random.randint(2, 3))
        for _, res in team_members.iterrows():
            s_date, e_date = get_random_date(-10, 90)
            allocations.append({
                'Project_ID': proj['Project_ID'],
                'Resource_ID': res['Resource_ID'],
                'Allocation_%': random.choice([0.5, 1.0]),
                'Start_Date': s_date,
                'End_Date': e_date
            })
    df_allocations = pd.DataFrame(allocations)

    # --- 5. DB_Pipeline (Future Demand) ---
    pipeline = []
    for i in range(1, 31):
        s_year = random.choice([2026, 2027])
        s_date = datetime.date(s_year, random.randint(1, 12), 1)
        e_date = s_date + relativedelta(months=random.randint(3, 9))
        
        pipeline.append({
            'Pipeline_ID': f'PIPE-{i:03d}',
            'Project_ID': f'NEW-PROJ-{i}',
            'Portfolio': random.choice(portfolios),
            'Team': random.choice(teams),
            'Goal': random.choice(quarters),
            'Skill_ID': random.choice(skills_data)['Skill_ID'],
            'Skill_Level_Needed': random.choice(['Standard', 'Senior']),
            'Start_Date': s_date,
            'End_Date': e_date,
            'Solution_Architect': 'TBD',
            'Product_Owner': 'TBD',
            'LTIM_Lead': 'TBD'
        })
    df_pipeline = pd.DataFrame(pipeline)

    # --- 6. DB_Financials (New!) ---
    financials = []
    for p in df_projects['Project_ID']:
        budget = random.randint(50, 500) * 1000
        actuals = budget * random.uniform(0.1, 0.6)
        financials.append({
            'Project_ID': p,
            'Total_Budget': budget,
            'Actuals_To_Date': int(actuals),
            'Forecast_To_Complete': int(budget - actuals),
            'Budget_Status': random.choice(['Green', 'Green', 'Amber'])
        })
    df_financials = pd.DataFrame(financials)

    # --- 7. Milestones, Updates, SLA ---
    milestones = []
    updates = []
    slas = []
    
    for p in df_projects['Project_ID']:
        # Milestone Logic
        base_d = datetime.date(2026, 4, 10)
        delay = random.choice([0, 0, 15])
        
        milestones.append({
            'Project_ID': p,
            'Milestone': 'Discovery',
            'Baseline_Date': datetime.date(2026, 2, 15),
            'Forecast_Date': datetime.date(2026, 2, 15),
            'Progress_Pct': 1.0, 
            'Status': 'Completed',
            'Comments': 'Done',
            'Risks_Issues': 'None'
        })
        milestones.append({
            'Project_ID': p,
            'Milestone': 'Build',
            'Baseline_Date': base_d,
            'Forecast_Date': base_d + datetime.timedelta(days=delay),
            'Progress_Pct': 0.4,
            'Status': 'Delayed' if delay > 0 else 'On Track',
            'Comments': 'In Progress',
            'Risks_Issues': 'None'
        })
        
        # Updates
        rag = random.choice(['Red', 'Amber', 'Green'])
        updates.append({
            'Project_ID': p,
            'Week': 'Wk 05',
            'RAG': rag,
            'Goal': 'Finalize UAT',
            'Narrative': 'Critical delay' if rag == 'Red' else 'Steady progress',
            'Tasks': '1. Task A\n2. Task B',
            'Risks': 'None' if rag == 'Green' else 'Staffing'
        })
        slas.append({'Project_ID': p, 'Metric': 'Defect SLA', 'Status': 'Met'})
        
    df_milestones = pd.DataFrame(milestones)
    df_updates = pd.DataFrame(updates)
    df_sla = pd.DataFrame(slas)

    return df_projects, df_resources, df_allocations, df_pipeline, df_skills, df_milestones, df_updates, df_sla, df_financials, df_config

# ==========================================
# 3. ENGINES (Demand Plan & Heatmap)
# ==========================================
def generate_demand_plan(df_pipe, df_skills):
    start_m = datetime.date(2026, 1, 1)
    months = get_month_columns(start_m, 24)
    
    unique_keys = df_pipe[['Portfolio', 'Team', 'Goal', 'Skill_ID', 'Skill_Level_Needed']].drop_duplicates()
    demand_rows = []
    
    for _, row in unique_keys.iterrows():
        sid = row['Skill_ID']
        s_name = df_skills[df_skills['Skill_ID'] == sid]['Skill_Name'].values[0]
        
        res_row = {
            'Portfolio': row['Portfolio'],
            'Team': row['Team'],
            'Goal': row['Goal'],
            'Skill Required': s_name,
            'Level': row['Skill_Level_Needed']
        }
        
        subset = df_pipe[
            (df_pipe['Portfolio'] == row['Portfolio']) & 
            (df_pipe['Team'] == row['Team']) & 
            (df_pipe['Goal'] == row['Goal']) &
            (df_pipe['Skill_ID'] == sid) & 
            (df_pipe['Skill_Level_Needed'] == row['Skill_Level_Needed'])
        ]
        
        for m in months:
            m_end = m + relativedelta(months=1, days=-1)
            count = 0
            for _, p in subset.iterrows():
                if p['Start_Date'] <= m_end and p['End_Date'] >= m:
                    count += 1
            res_row[m] = count
            
        demand_rows.append(res_row)
    return pd.DataFrame(demand_rows), months

def generate_heatmap_data(df_res, df_alloc, df_skills):
    today = datetime.date.today().replace(day=1)
    months = get_month_columns(today, 12)
    heatmap_rows = []
    
    for _, res in df_res.iterrows():
        sid = res['Skill_ID']
        s_series = df_skills[df_skills['Skill_ID'] == sid]['Skill_Name']
        s_name = s_series.values[0] if not s_series.empty else sid
        
        row = {'Resource Name': res['Full_Name'], 'Primary Skill': s_name, 'Manager': res['Manager']}
        my_allocs = df_alloc[df_alloc['Resource_ID'] == res['Resource_ID']]
        
        for m in months:
            load = 0
            m_end = m + relativedelta(months=1, days=-1)
            for _, a in my_allocs.iterrows():
                if a['Start_Date'] <= m_end and a['End_Date'] >= m:
                    load += a['Allocation_%']
            row[m] = load
        heatmap_rows.append(row)
    return pd.DataFrame(heatmap_rows), months

# ==========================================
# 4. EXCEL ORCHESTRATION
# ==========================================
def create_workbook_v4():
    dfs = create_database_v4()
    df_p, df_r, df_a, df_pipe, df_s, df_m, df_u, df_sla, df_fin, df_config = dfs
    
    df_demand, _ = generate_demand_plan(df_pipe, df_s)
    df_heat, _ = generate_heatmap_data(df_r, df_a, df_s)
    
    writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
    wb = writer.book

    # --- FORMATS ---
    f_navy = wb.add_format({'bold': True, 'fg_color': '#0F2C4C', 'font_color': 'white', 'border': 1, 'align': 'center'})
    f_grey = wb.add_format({'bold': True, 'fg_color': '#444444', 'font_color': 'white', 'border': 1, 'align': 'center'})
    f_rich = wb.add_format({'text_wrap': True, 'valign': 'top', 'border': 1})
    f_cen = wb.add_format({'align': 'center', 'valign': 'top', 'border': 1})
    f_money = wb.add_format({'num_format': '$#,##0', 'align': 'center', 'border': 1})
    
    rag_red = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True, 'align': 'center', 'border': 1})
    rag_grn = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': True, 'align': 'center', 'border': 1})
    rag_amb = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'bold': True, 'align': 'center', 'border': 1})
    
    # --- 0. NAVIGATION / HOME TAB ---
    ws_nav = wb.add_worksheet(">> HOME <<")
    ws_nav.hide_gridlines(2)
    ws_nav.write('B2', "PORTFOLIO MANAGEMENT SYSTEM v4.0", wb.add_format({'bold': True, 'font_size': 20, 'font_color': '#0F2C4C'}))
    
    nav_buttons = [
        ('>> DASHBOARD <<', 'View Executive Report'),
        ('>> DEMAND_PLAN <<', 'Future Resource Needs'),
        ('>> RES_HEATMAP <<', 'Current Capacity'),
        ('DB_Projects', 'Manage Projects'),
        ('DB_Allocations', 'Assign Resources'),
        ('DB_Milestones', 'Track Milestones'),
        ('DB_Financials', 'Budget vs Actuals'),
        ('DB_Resources', 'Resource Master'),
        ('DB_Pipeline', 'Future Opportunities')
    ]
    
    row = 4
    for sheet, desc in nav_buttons:
        # Create "Button" look with hyperlink
        ws_nav.write_url(row, 1, f"internal:'{sheet}'!A1", string=sheet, cell_format=wb.add_format({'bold': True, 'font_color': 'blue', 'underline': True, 'font_size': 12}))
        ws_nav.write(row, 2, desc)
        row += 1

    # --- 1. DASHBOARD ---
    # Aggregate Data
    dash_rows = []
    total_budget = 0
    total_spent = 0
    
    for _, p in df_p.iterrows():
        pid = p['Project_ID']
        upd = df_u[df_u['Project_ID'] == pid].iloc[0]
        fin = df_fin[df_fin['Project_ID'] == pid].iloc[0]
        
        total_budget += fin['Total_Budget']
        total_spent += fin['Actuals_To_Date']
        
        # Roadmap Logic
        m_rows = df_m[df_m['Project_ID'] == pid]
        m_lines = []
        for _, m in m_rows.iterrows():
            d = (m['Forecast_Date'] - m['Baseline_Date']).days
            icon = "✅" if m['Status'] == 'Completed' else ("⚠️" if d > 0 else "🔵")
            m_lines.append(f"{icon} {m['Milestone']} ({int(m['Progress_Pct']*100)}%)")
        
        # Resource Logic
        allocs = df_a[df_a['Project_ID'] == pid].merge(df_r, on='Resource_ID').merge(df_s, on='Skill_ID')
        t_lines = [f"• {r['Full_Name']} ({r['Skill_Name']})" for _, r in allocs.iterrows()]
        
        dash_rows.append({
            'Project': p['Project_Name'], 'Portfolio': p['Portfolio'], 'Team': p['Team'], 'Goal': p['Goal'],
            'Status': upd['RAG'], 'Budget_Status': fin['Budget_Status'], 
            'Roadmap': "\n".join(m_lines), 'Resources': "\n".join(t_lines),
            'Narrative': f"GOAL: {upd['Goal']}\nRISK: {upd['Risks']}"
        })
        
    df_dash = pd.DataFrame(dash_rows).sort_values(['Portfolio', 'Team'])
    
    ws_dash = wb.add_worksheet(">> DASHBOARD <<")
    
    # EXECUTIVE SUMMARY HEADER
    ws_dash.merge_range('A1:C1', "EXECUTIVE SUMMARY", f_navy)
    ws_dash.write('A2', "Total Projects:", f_cen)
    ws_dash.write('B2', len(df_dash), f_rich)
    ws_dash.write('A3', "Total Budget:", f_cen)
    ws_dash.write('B3', total_budget, f_money)
    ws_dash.write('A4', "Budget Utilized:", f_cen)
    ws_dash.write('B4', total_spent / total_budget if total_budget else 0, wb.add_format({'num_format': '0%', 'border': 1}))
    
    # Table Header
    start_row = 6
    cols = ['Project', 'Portfolio', 'Team', 'Goal', 'Status', 'Budget_Status', 'Roadmap', 'Resources', 'Narrative']
    ws_dash.write_row(start_row, 0, cols, f_navy)
    
    for idx, row in enumerate(df_dash.itertuples(index=False), start_row + 1):
        ws_dash.write(idx, 0, row.Project, f_rich)
        ws_dash.write(idx, 1, row.Portfolio, f_cen)
        ws_dash.write(idx, 2, row.Team, f_cen)
        ws_dash.write(idx, 3, row.Goal, f_cen)
        
        ws_dash.write(idx, 4, row.Status, rag_grn if row.Status=='Green' else (rag_red if row.Status=='Red' else rag_amb))
        ws_dash.write(idx, 5, row.Budget_Status, rag_grn if row.Budget_Status=='Green' else rag_amb)
        
        ws_dash.write(idx, 6, row.Roadmap, f_rich)
        ws_dash.write(idx, 7, row.Resources, f_rich)
        ws_dash.write(idx, 8, row.Narrative, f_rich)
        
    ws_dash.set_column('A:A', 25)
    ws_dash.set_column('G:I', 30)

    # --- 2. DEMAND PLAN & HEATMAP (Standard) ---
    ws_dem = wb.add_worksheet(">> DEMAND_PLAN <<")
    df_demand.to_excel(writer, sheet_name=">> DEMAND_PLAN <<", index=False) # Simplified for brevity, add formatting as needed
    
    ws_heat = wb.add_worksheet(">> RES_HEATMAP <<")
    df_heat.to_excel(writer, sheet_name=">> RES_HEATMAP <<", index=False)

    # --- 3. DB TABS (With Tables & Validation) ---
    
    # Helper to add table
    def add_db_sheet(df, name):
        df.to_excel(writer, sheet_name=name, index=False)
        ws = writer.sheets[name]
        (max_row, max_col) = df.shape
        options = {'columns': [{'header': col} for col in df.columns]}
        ws.add_table(0, 0, max_row, max_col - 1, options)
        return ws
        
    # Config (Hidden)
    df_config.to_excel(writer, sheet_name="DB_Config", index=False)
    writer.sheets['DB_Config'].hide()
    # Define Named Ranges for Validation
    wb.define_name('List_Goals', '=DB_Config!$A$2:$A$33')
    
    # Write DBs
    ws_proj = add_db_sheet(df_p, "DB_Projects")
    ws_res = add_db_sheet(df_r, "DB_Resources")
    ws_skill = add_db_sheet(df_s, "DB_Skills")
    ws_fin = add_db_sheet(df_fin, "DB_Financials")
    
    # Define Named Ranges for IDs
    wb.define_name('List_ProjectIDs', f'=DB_Projects!$A$2:$A${len(df_p)+1}')
    wb.define_name('List_ResourceIDs', f'=DB_Resources!$A$2:$A${len(df_r)+1}')
    wb.define_name('List_SkillIDs', f'=DB_Skills!$A$2:$A${len(df_s)+1}')

    # Allocations (With Validation)
    ws_alloc = add_db_sheet(df_a, "DB_Allocations")
    ws_alloc.data_validation(f'A2:A{len(df_a)+50}', {'validate': 'list', 'source': '=List_ProjectIDs'})
    ws_alloc.data_validation(f'B2:B{len(df_a)+50}', {'validate': 'list', 'source': '=List_ResourceIDs'})
    
    # Pipeline (With Goal Validation)
    ws_pipe = add_db_sheet(df_pipe, "DB_Pipeline")
    ws_pipe.data_validation(f'E2:E{len(df_pipe)+50}', {'validate': 'list', 'source': '=List_Goals'})

    # Others
    add_db_sheet(df_m, "DB_Milestones")
    add_db_sheet(df_u, "DB_Updates")
    add_db_sheet(df_sla, "DB_SLA")

    writer.close()
    print(f"✅ Generated v4 System: {OUTPUT_FILE}")

if __name__ == "__main__":
    create_workbook_v4()