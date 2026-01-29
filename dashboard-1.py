import pandas as pd
import random
import datetime
from xlsxwriter.utility import xl_rowcol_to_cell

# ==========================================
# CONFIGURATION
# ==========================================
OUTPUT_FILE = 'Dynamic_Weekly_Report.xlsx'

# ==========================================
# 1. GENERATE SOURCE DATA (MOCK DATABASE)
# ==========================================

def create_mock_database():
    # --- 1. Project Master ---
    projects = []
    for i in range(1, 21):
        projects.append({
            'Project_ID': f'P{i:03d}',
            'Project_Name': f'Project {chr(65+(i%26))}-{i} (Stream {i%3+1})',
            'Lead': random.choice(['Suresh', 'Abhi', 'Brandi', 'Stanley']),
            'PM': random.choice(['Kalpesh', 'Rail', 'Sigmas']),
            'Kickoff_Date': (datetime.date.today() - datetime.timedelta(days=random.randint(20, 100))).strftime('%Y-%m-%d'),
            'End_Date': (datetime.date.today() + datetime.timedelta(days=random.randint(60, 200))).strftime('%Y-%m-%d')
        })
    df_projects = pd.DataFrame(projects)

    # --- 2. Resource Master ---
    resources = []
    names = ['Mike', 'Sarah', 'Jessica', 'Tom', 'Ravi', 'Siva', 'Wei', 'Lisa']
    roles = ['Arch', 'Dev', 'QA', 'BA', 'Lead']
    for idx, name in enumerate(names):
        resources.append({
            'Resource_ID': f'R{idx:03d}',
            'Name': name,
            'Role': roles[idx % len(roles)]
        })
    df_resources = pd.DataFrame(resources)

    # --- 3. Allocations (Many-to-Many) ---
    allocations = []
    for p in df_projects['Project_ID']:
        # Assign 2-4 random resources per project
        team = random.sample(resources, k=random.randint(2, 4))
        for res in team:
            allocations.append({
                'Project_ID': p,
                'Resource_ID': res['Resource_ID'],
                'Allocation_Pct': random.choice([0.25, 0.50, 1.0])
            })
    df_allocations = pd.DataFrame(allocations)

    # --- 4. Milestones (One-to-Many) ---
    milestones = []
    phases = ['Discovery', 'Design', 'Build', 'UAT', 'Go-Live']
    for p in df_projects['Project_ID']:
        for idx, phase in enumerate(phases):
            # Stagger dates
            date = datetime.date.today() + datetime.timedelta(days=(idx*20) - 20)
            status = 'Completed' if idx < 2 else 'In Progress' if idx == 2 else 'Pending'
            milestones.append({
                'Project_ID': p,
                'Milestone_Name': phase,
                'Due_Date': date.strftime('%Y-%m-%d'),
                'Status': status,
                'Completion_Pct': 1.0 if status == 'Completed' else (0.5 if status == 'In Progress' else 0)
            })
    df_milestones = pd.DataFrame(milestones)

    # --- 5. Weekly Status (One-to-One) ---
    updates = []
    for p in df_projects['Project_ID']:
        status = random.choice(['Red', 'Amber', 'Green'])
        updates.append({
            'Project_ID': p,
            'Week_No': 'Wk 04',
            'Overall_RAG': status,
            'Goal_This_Week': 'Finalize UAT Sign-off and prepare deployment scripts.',
            'Key_Achievements': 'Completed Module B coding. All Priority 1 bugs closed.',
            'Top_Risk': 'Data latency issues in QA env' if status == 'Red' else 'None',
            'Weekly_Status_Narrative': 'Progress is steady. ' + ('Critical blocker on API.' if status == 'Red' else 'On track.')
        })
    df_updates = pd.DataFrame(updates)

    # --- 6. SLA (One-to-Many) ---
    slas = []
    for p in df_projects['Project_ID']:
        slas.append({'Project_ID': p, 'Metric': 'Defect Turnaround', 'Status': random.choice(['Met', 'Met', 'Breached'])})
        slas.append({'Project_ID': p, 'Metric': 'Uptime', 'Status': 'Met'})
    df_sla = pd.DataFrame(slas)

    return df_projects, df_resources, df_allocations, df_milestones, df_updates, df_sla

# ==========================================
# 2. LOGIC: AGGREGATE DATA FOR DASHBOARD
# ==========================================
def compile_dashboard_data(dfs):
    df_p, df_r, df_a, df_m, df_u, df_s = dfs
    
    dashboard_rows = []

    for _, proj in df_p.iterrows():
        pid = proj['Project_ID']
        
        # 1. Get Weekly Update
        update = df_u[df_u['Project_ID'] == pid].iloc[0]
        
        # 2. Aggregate Milestones (The "Rich Text" Logic)
        proj_miles = df_m[df_m['Project_ID'] == pid].sort_values('Due_Date')
        mile_lines = []
        for _, m in proj_miles.iterrows():
            icon = "✅" if m['Status'] == 'Completed' else "⚠️" if m['Status'] == 'In Progress' else "⚪"
            date_str = m['Due_Date'][5:] # mm-dd
            mile_lines.append(f"{icon} {m['Milestone_Name']} ({int(m['Completion_Pct']*100)}%) - {date_str}")
        roadmap_str = "\n".join(mile_lines)
        
        # 3. Aggregate Resources (Join Allocation + Resource Master)
        proj_alloc = df_a[df_a['Project_ID'] == pid].merge(df_r, on='Resource_ID')
        team_lines = []
        total_util = 0
        for _, r in proj_alloc.iterrows():
            pct = r['Allocation_Pct']
            team_lines.append(f"• {r['Name']} ({r['Role']}): {int(pct*100)}%")
            total_util += pct
        team_str = "\n".join(team_lines)
        
        # 4. Aggregate SLA
        proj_sla = df_s[df_s['Project_ID'] == pid]
        sla_status = "Breached" if "Breached" in proj_sla['Status'].values else "Met"

        # Build Row
        dashboard_rows.append({
            'Project Name': proj['Project_Name'],
            'Lead / PM': f"Lead: {proj['Lead']}\nPM: {proj['PM']}",
            'Overall Status': update['Overall_RAG'],
            'Milestone Roadmap': roadmap_str,
            'Resource Plan': team_str,
            'Total Capacity': total_util, # Sum of allocations (e.g., 2.5 FTE)
            'Weekly Update': f"GOAL: {update['Goal_This_Week']}\n\nUPDATE: {update['Weekly_Status_Narrative']}\n\nRISK: {update['Top_Risk']}",
            'SLA': sla_status
        })
        
    return pd.DataFrame(dashboard_rows)

# ==========================================
# 3. GENERATE EXCEL FILE
# ==========================================
def create_full_workbook():
    # Get Data
    dfs = create_mock_database()
    df_dashboard = compile_dashboard_data(dfs)
    
    writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
    workbook = writer.book

    # --- FORMATS ---
    header_fmt = workbook.add_format({'bold': True, 'fg_color': '#0F2C4C', 'font_color': 'white', 'border': 1, 'valign': 'vcenter', 'align': 'center'})
    
    # The "Rich" Format (Text Wrap is Crucial)
    rich_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top', 'border': 1, 'font_size': 10})
    
    center_fmt = workbook.add_format({'align': 'center', 'valign': 'top', 'border': 1})
    
    # RAG Formats
    rag_red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'align': 'center', 'border': 1, 'bold': True})
    rag_amber = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'align': 'center', 'border': 1, 'bold': True})
    rag_green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'align': 'center', 'border': 1, 'bold': True})

    # ================= 1. THE VIEW (DASHBOARD) =================
    sheet_name = '>> REPORT <<'
    df_dashboard.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
    ws = writer.sheets[sheet_name]
    
    # Title
    ws.merge_range('A1:H1', "WEEKLY EXECUTIVE STATUS REPORT (AUTO-GENERATED)", workbook.add_format({'bold': True, 'font_size': 16, 'fg_color': '#EFEFEF', 'align': 'center'}))
    
    # Write Headers
    for col_num, value in enumerate(df_dashboard.columns.values):
        ws.write(1, col_num, value, header_fmt)
        
    # Column Config
    ws.set_column('A:A', 25, rich_fmt) # Project Name
    ws.set_column('B:B', 15, rich_fmt) # Lead/PM
    ws.set_column('C:C', 10, center_fmt) # Status
    ws.set_column('D:D', 30, rich_fmt) # Roadmap (Wide)
    ws.set_column('E:E', 25, rich_fmt) # Resources
    ws.set_column('F:F', 10, center_fmt) # Capacity
    ws.set_column('G:G', 40, rich_fmt) # Narrative (Very Wide)
    ws.set_column('H:H', 10, center_fmt) # SLA
    
    # Conditional Formatting
    ws.conditional_format('C3:C100', {'type': 'cell', 'criteria': 'equal to', 'value': '"Red"', 'format': rag_red})
    ws.conditional_format('C3:C100', {'type': 'cell', 'criteria': 'equal to', 'value': '"Amber"', 'format': rag_amber})
    ws.conditional_format('C3:C100', {'type': 'cell', 'criteria': 'equal to', 'value': '"Green"', 'format': rag_green})
    
    ws.conditional_format('H3:H100', {'type': 'cell', 'criteria': 'equal to', 'value': '"Breached"', 'format': rag_red})

    # ================= 2. THE SOURCE TABS (DATABASE) =================
    # Write the raw data tabs so the user can see "Under the hood"
    sheet_names = ['DB_Projects', 'DB_Resources', 'DB_Allocations', 'DB_Milestones', 'DB_Updates', 'DB_SLA']
    for i, df in enumerate(dfs):
        s_name = sheet_names[i]
        df.to_excel(writer, sheet_name=s_name, index=False)
        # Auto-width columns roughly
        writer.sheets[s_name].set_column(0, len(df.columns)-1, 15)

    writer.close()
    print(f"✅ Generated Dynamic Workbook: {OUTPUT_FILE}")

if __name__ == "__main__":
    create_full_workbook()