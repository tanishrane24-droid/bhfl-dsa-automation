"""
Demo Data Generator for BHFL DSA Automation
Generates sample Master.xlsx, DSA_MAIL_MAP.xlsx, and demo files
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import random

# Create directories
BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "Input"
CONFIG_DIR = BASE_DIR / "config"
DEMO_DIR = BASE_DIR / "Demo_Output"
LOGS_DIR = BASE_DIR / "Logs"

for dir_path in [INPUT_DIR, CONFIG_DIR, DEMO_DIR, LOGS_DIR]:
    dir_path.mkdir(exist_ok=True)

print("=" * 80)
print("BHFL DSA AUTOMATION - DEMO DATA GENERATOR")
print("=" * 80)

# ============================================================================
# 1. GENERATE MASTER.XLSX WITH 100 SAMPLE ROWS
# ============================================================================

print("\n[1/5] Generating Master.xlsx with 100 sample rows...")

DSA_NAMES = [
    "ABC Finance",
    "XYZ Associates",
    "PQR Services",
    "Prime Housing",
    "Elite Partners",
    "Capital Connect",
    "Sai Associates",
    "Home Link Finance",
    "Urban Capital",
    "Maharashtra DSA"
]

LOCATIONS = ["Mumbai", "Pune", "Nashik", "Nagpur", "Thane"]

DSM_IDS = ["DSM1001", "DSM1002", "DSM1003", "DSM1004", "DSM1005"]

RSM_NAMES = ["RSM Mumbai", "RSM West", "RSM Central"]

RM_NAMES = ["RM1", "RM2", "RM3", "RM4"]

SM_NAMES = ["SM_A", "SM_B", "SM_C"]

BH_NAMES = ["BH_Mumbai", "BH_Pune", "BH_Nashik"]

# Generate 100 rows with repeated DSA entries
master_data = []
records_per_dsa = 10  # 10 DSAs * 10 records = 100 rows

for dsa in DSA_NAMES:
    for i in range(records_per_dsa):
        row = {
            'DSM ID': random.choice(DSM_IDS),
            'RSM Name': random.choice(RSM_NAMES),
            'PAN No': f"AAAA{random.randint(1000, 9999)}A",
            'DSA Name': dsa,
            'Location': random.choice(LOCATIONS),
            'RM': random.choice(RM_NAMES),
            'SM': random.choice(SM_NAMES),
            'BH': random.choice(BH_NAMES)
        }
        master_data.append(row)

master_df = pd.DataFrame(master_data)
master_file = INPUT_DIR / "Master.xlsx"
master_df.to_excel(master_file, index=False, sheet_name='Data')
print(f"✓ Created {master_file.name} with {len(master_df)} rows")
print(f"  - 10 DSA Names (10 records each)")
print(f"  - Sample DSAs: {', '.join(DSA_NAMES[:3])}...")

# ============================================================================
# 2. GENERATE DSA_MAIL_MAP.XLSX
# ============================================================================

print("\n[2/5] Generating DSA_MAIL_MAP.xlsx...")

mail_map_data = []
for dsa in DSA_NAMES:
    # Generate demo emails
    dsa_clean = dsa.lower().replace(' ', '_')
    email = f"{dsa_clean}.demo@test.com"
    mail_map_data.append({
        'DSA Name': dsa,
        'Mail ID': email
    })

mail_map_df = pd.DataFrame(mail_map_data)
mail_map_file = CONFIG_DIR / "DSA_MAIL_MAP.xlsx"
mail_map_df.to_excel(mail_map_file, index=False, sheet_name='Map')
print(f"✓ Created {mail_map_file.name}")
print(f"  - 10 DSA Email Mappings")
for idx, row in mail_map_df.iterrows():
    print(f"    {row['DSA Name']:20} → {row['Mail ID']}")

# ============================================================================
# 3. GENERATE SAMPLE RUN_STATUS.XLSX
# ============================================================================

print("\n[3/5] Generating Sample RUN_STATUS.xlsx...")

status_data = []
for dsa in DSA_NAMES:
    email = f"{dsa.lower().replace(' ', '_')}.demo@test.com"
    status_data.append({
        'Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'DSA': dsa,
        'Mail': email,
        'Status': 'SUCCESS',
        'Remarks': 'Mail sent successfully (Demo)'
    })

status_df = pd.DataFrame(status_data)
status_file = LOGS_DIR / "RUN_STATUS_DEMO.xlsx"
status_df.to_excel(status_file, index=False, sheet_name='Status')
print(f"✓ Created {status_file.name}")
print(f"  - 10 Status Records (All SUCCESS)")

# ============================================================================
# 4. GENERATE DEMO MAIL FILES
# ============================================================================

print("\n[4/5] Generating Demo Mail Files...")

# Create individual mail files for each DSA
for dsa in DSA_NAMES:
    dsa_clean = dsa.lower().replace(' ', '_')
    email = f"{dsa_clean}.demo@test.com"
    
    # Create mail content
    mail_content = f"""================================================================================
DEMONSTRATION EMAIL - NOT ACTUALLY SENT
================================================================================

TO: {email}
SUBJECT: DSA Data Distribution - {dsa}
DATE: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

================================================================================

Dear {dsa},

Please find attached the data distribution file for your records.

File: {dsa.replace(' ', '_')}.xlsx
Records: {records_per_dsa}
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

This is a demonstration email showing what would be sent in production mode.

Data Summary:
- DSA Name: {dsa}
- Total Records: {records_per_dsa}
- Locations: {', '.join(LOCATIONS)}
- Report Period: {datetime.now().strftime('%B %Y')}

ATTACHMENT: {dsa.replace(' ', '_')}.xlsx

Best regards,
BHFL DSA Automation System

================================================================================
NOTE: This is a DEMO MODE email. In production mode, actual mails are sent via
Outlook with real attachments.
================================================================================
\n\"\"\"\n\n    # Save mail file\n    mail_file = DEMO_DIR / f\"Mail_{dsa_clean.upper()}.txt\"\n    with open(mail_file, 'w') as f:\n        f.write(mail_content)\n    print(f\"✓ Created {mail_file.name}\")\n\n# ============================================================================\n# 5. GENERATE SAMPLE LOGS\n# ============================================================================\n\nprint(\"\\n[5/5] Generating Sample Log Files...\")\n\n# Error log (sample)\nerror_data = []\nerror_log_file = LOGS_DIR / \"ERROR_LOG_DEMO.xlsx\"\nerror_df = pd.DataFrame({\n    'Timestamp': [],\n    'Error Type': [],\n    'DSA': [],\n    'Details': []\n})\nerror_df.to_excel(error_log_file, index=False, sheet_name='Errors')\nprint(f\"✓ Created {error_log_file.name} (Sample - No errors in demo)\")\n\n# Detailed execution log\nexecution_log_file = LOGS_DIR / \"automation_demo.log\"\nlog_content = f\"\"\"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ============================================================\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] BHFL DSA AUTOMATION - DEMO RUN\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ============================================================\n\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [1/6] Reading Master.xlsx...\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ✓ Read 100 records from Master.xlsx\n\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [2/6] Identifying unique DSA names...\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ✓ Detected 10 unique DSAs:\n\"\"\"\n\nfor dsa in DSA_NAMES:\n    log_content += f\"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]   - {dsa}\\n\"\n\nlog_content += f\"\"\"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [3/6] Creating DSA-wise Excel files...\n\"\"\"\n\nfor dsa in DSA_NAMES:\n    dsa_clean = dsa.replace(' ', '_')\n    log_content += f\"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ✓ Created {dsa_clean}.xlsx ({records_per_dsa} records)\\n\"\n\nlog_content += f\"\"\"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [4/6] Loading DSA mail map...\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ✓ Loaded mail map for 10 DSAs\n\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [5/6] Sending Outlook mails with attachments...\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] *** DEMO MODE: Mails NOT actually sent ***\n\"\"\"\n\nfor dsa in DSA_NAMES:\n    dsa_clean = dsa.lower().replace(' ', '_')\n    email = f\"{dsa_clean}.demo@test.com\"\n    log_content += f\"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ✓ Mail prepared for {dsa} ({email}) - NOT SENT (DEMO MODE)\\n\"\n\nlog_content += f\"\"\"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [6/6] Archiving processed files...\n\"\"\"\n\nfor dsa in DSA_NAMES:\n    dsa_clean = dsa.replace(' ', '_')\n    log_content += f\"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ✓ Archived {dsa_clean}.xlsx\\n\"\n\nlog_content += f\"\"\"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ============================================================\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] AUTOMATION COMPLETED SUCCESSFULLY (DEMO MODE)\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ============================================================\n\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] SUMMARY DASHBOARD:\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]   Total DSAs:      10\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]   Mails Prepared:  10\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]   Failed:          0\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]   Success Rate:    100.0%\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] NOTE: This is a DEMO RUN. No actual emails were sent.\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ============================================================\n\"\"\"\n\nwith open(execution_log_file, 'w') as f:\n    f.write(log_content)\nprint(f\"✓ Created {execution_log_file.name}\")\n\n# ============================================================================\n# 6. CREATE SAMPLE OUTPUT FILES\n# ============================================================================\n\nprint(\"\\n[6/6] Creating Sample Output Files for Demo...\")\n\n# Create sample DSA files in Demo_Output\nfor dsa in DSA_NAMES:\n    dsa_clean = dsa.replace(' ', '_')\n    \n    # Filter master data for this DSA\n    dsa_data = master_df[master_df['DSA Name'] == dsa]\n    \n    # Save to demo output\n    demo_file = DEMO_DIR / f\"{dsa_clean}.xlsx\"\n    dsa_data.to_excel(demo_file, index=False, sheet_name='Data')\n    print(f\"✓ Created {demo_file.name} ({len(dsa_data)} records)\")\n\n# ============================================================================\n# SUMMARY\n# ============================================================================\n\nprint(\"\\n\" + \"=\" * 80)\nprint(\"DEMO DATA GENERATION COMPLETED SUCCESSFULLY!\")\nprint(\"=\" * 80)\n\nprint(\"\\n📁 FILES GENERATED:\\n\")\n\nprint(\"INPUT DATA:\")\nprint(f\"  ✓ Input/Master.xlsx\")\nprint(f\"    - 100 sample rows\")\nprint(f\"    - 10 DSA Names (10 records each)\")\nprint(f\"    - Realistic sample data\")\n\nprint(\"\\nCONFIGURATION:\")\nprint(f\"  ✓ config/DSA_MAIL_MAP.xlsx\")\nprint(f\"    - 10 DSA Email mappings\")\nprint(f\"    - Demo email addresses\")\n\nprint(\"\\nLOGS:\")\nprint(f\"  ✓ Logs/RUN_STATUS_DEMO.xlsx\")\nprint(f\"    - Sample status report\")\nprint(f\"  ✓ Logs/ERROR_LOG_DEMO.xlsx\")\nprint(f\"    - Error tracking (empty - no errors in demo)\")\nprint(f\"  ✓ Logs/automation_demo.log\")\nprint(f\"    - Detailed execution log\")\n\nprint(\"\\nDEMO EMAILS:\")\nprint(f\"  ✓ Demo_Output/ folder\")\nprint(f\"    - 10 sample mail files\")\nprint(f\"    - Shows what would be sent to each DSA\")\n\nprint(\"\\nSAMPLE OUTPUT:\")\nprint(f\"  ✓ Demo_Output/ folder\")\nprint(f\"    - 10 DSA Excel files\")\nprint(f\"    - Shows generated output\")\n\nprint(\"\\n\" + \"=\" * 80)\nprint(\"NEXT STEPS:\")\nprint(\"=\" * 80)\nprint(\"\\n1. Run the application:\")\nprint(\"   python main.py\")\n\nprint(\"\\n2. Click 'RUN AUTOMATION' button\")\n\nprint(\"\\n3. Application will:\")\nprint(\"   - Auto-detect Master.xlsx\")\nprint(\"   - Auto-detect DSA_MAIL_MAP.xlsx\")\nprint(\"   - Process all data\")\nprint(\"   - Show live progress\")\n\nprint(\"\\n4. View results in:\")\nprint(f\"   - Output/ folder (generated files)\")\nprint(f\"   - Archive/ folder (processed files)\")\nprint(f\"   - Logs/ folder (status reports)\")\n\nprint(\"\\n\" + \"=\" * 80)\nprint(\"READY FOR MANAGER DEMO!\")\nprint(\"=\" * 80 + \"\\n\")\n"