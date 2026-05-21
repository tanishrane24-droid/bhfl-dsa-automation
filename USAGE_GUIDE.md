# BHFL DSA Automation - User Guide

## Quick Start (1 Minute)

### Step 1: Prepare Files

1. **Input/Master.xlsx** - Your data file with DSA column
2. **config/DSA_MAIL_MAP.xlsx** - Email mapping for each DSA

### Step 2: Run Application

Double-click: `BHFL_DSA_AUTOMATION.exe`

### Step 3: Click ONE Button

Click: **"▶ RUN AUTOMATION"**

### Done!

Everything runs automatically:
- ✓ Reads your data
- ✓ Splits by DSA
- ✓ Creates Excel files
- ✓ Sends via Outlook
- ✓ Archives files
- ✓ Generates reports

---

## Detailed Workflow

### Application Window

```
╔════════════════════════════════════════════════════════════╗
║           BHFL DSA AUTOMATION                               ║
║        One-click DSA Mail Distribution System               ║
╠════════════════════════════════════════════════════════════╣
║ Status                                                       ║
║ ✓ Master File: Master.xlsx                                  ║
║ ✓ Mail Map: DSA_MAIL_MAP.xlsx                               ║
╠════════════════════════════════════════════════════════════╣
║ Progress                                                     ║
║ [████████░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░] 40%        ║
║ Creating DSA-wise files...                                   ║
╠════════════════════════════════════════════════════════════╣
║ Live Console Output                                          ║
║ ✓ Master file detected: Input/Master.xlsx                   ║
║ ✓ Mail map detected: config/DSA_MAIL_MAP.xlsx               ║
║ ✓ Read 500 records from Master.xlsx                         ║
║ ✓ Detected 5 unique DSAs:                                   ║
║   - ABC Finance                                              ║
║   - XYZ Associates                                           ║
║   - QRS Brokers                                              ║
║   - ... (more DSAs)                                          ║
╠════════════════════════════════════════════════════════════╣
║ [▶ RUN AUTOMATION] [📁 OPEN OUTPUT] [📋 VIEW LOGS] [❌ EXIT] ║
╚════════════════════════════════════════════════════════════╝
```

---

## What Each Button Does

### ▶ RUN AUTOMATION
- Starts the entire process
- No user interaction needed
- All steps run automatically
- Takes 2-5 minutes depending on data size

### 📁 OPEN OUTPUT
- Opens the Output folder
- Shows generated DSA files (before archiving)
- Shows archived files after completion

### 📋 VIEW LOGS
- Opens the Logs folder
- Shows detailed execution logs
- Shows status reports (Excel format)

### ❌ EXIT
- Closes the application
- Does NOT interrupt running automation

---

## Automation Steps (Automatic)

### Step 1: Read Master File
```
✓ Master file detected: Input/Master.xlsx
✓ Read 500 records from Master.xlsx
```
- Application reads your input data
- No action needed

### Step 2: Detect DSA Names
```
✓ Detected 5 unique DSAs:
  - ABC Finance
  - XYZ Associates
  - QRS Brokers
  - PQR Consultants
  - STU Partners
```
- Application identifies all unique DSA names
- No action needed

### Step 3: Create DSA Files
```
✓ Created ABC_Finance.xlsx (100 records)
✓ Created XYZ_Associates.xlsx (120 records)
✓ Created QRS_Brokers.xlsx (80 records)
✓ Created PQR_Consultants.xlsx (110 records)
✓ Created STU_Partners.xlsx (90 records)
```
- Each DSA gets its own Excel file
- Files created in Output folder
- No action needed

### Step 4: Load Mail Configuration
```
✓ Loaded mail map for 5 DSAs
```
- Application reads email addresses from config
- No action needed

### Step 5: Send Mails
```
✓ Mail sent to ABC Finance (abc.finance@company.com)
✓ Mail sent to XYZ Associates (xyz.associates@company.com)
✓ Mail sent to QRS Brokers (qrs.brokers@company.com)
✓ Mail sent to PQR Consultants (pqr.consultants@company.com)
✓ Mail sent to STU Partners (stu.partners@company.com)
```
- Outlook automatically creates and sends emails
- Files automatically attached
- No confirmation popups
- No manual actions
- Retry logic: If mail fails, tries 3 times automatically

### Step 6: Archive Files
```
✓ Archived ABC_Finance.xlsx
✓ Archived XYZ_Associates.xlsx
✓ Archived QRS_Brokers.xlsx
✓ Archived PQR_Consultants.xlsx
✓ Archived STU_Partners.xlsx
```
- Generated files moved to Archive folder
- Keeps Output folder clean for next run
- No action needed

### Final Report
```
============================================================
AUTOMATION COMPLETED SUCCESSFULLY
============================================================

SUMMARY DASHBOARD:
  Total DSAs:      5
  Mails Sent:      5
  Failed:          0
  Success Rate:    100.0%

============================================================
```

---

## Popup After Completion

```
╔══════════════════════════════════════╗
║          SUCCESS                     ║
║                                      ║
║  AUTOMATION COMPLETED                ║
║                                      ║
║  Total DSAs: 5                       ║
║  Mails Sent: 5                       ║
║  Failed: 0                           ║
║  Success Rate: 100.0%                ║
║                                      ║
║              [OK]                    ║
╚══════════════════════════════════════╝
```

---

## Output Files

### After Automation Completes:

#### 1. Generated DSA Files (Archived)
```
Archive/
├── ABC_Finance.xlsx
├── XYZ_Associates.xlsx
├── QRS_Brokers.xlsx
├── PQR_Consultants.xlsx
└── STU_Partners.xlsx
```

#### 2. Status Report
```
Logs/RUN_STATUS_20240521_143025.xlsx

Columns:
Date              | DSA               | Mail                     | Status  | Remarks
2024-05-21 14:30  | ABC Finance       | abc.finance@company.com  | SUCCESS | Mail sent successfully
2024-05-21 14:31  | XYZ Associates    | xyz@company.com          | SUCCESS | Mail sent successfully
2024-05-21 14:32  | QRS Brokers       | qrs@company.com          | SUCCESS | Mail sent successfully
2024-05-21 14:33  | PQR Consultants   | pqr@company.com          | SUCCESS | Mail sent successfully
2024-05-21 14:34  | STU Partners      | stu@company.com          | SUCCESS | Mail sent successfully
```

#### 3. Detailed Logs
```
Logs/automation_20240521_143025.log

[Full execution log with timestamps]
```

---

## Common Scenarios

### Scenario 1: Quick Run
```
1. Double-click BHFL_DSA_AUTOMATION.exe
2. Click "▶ RUN AUTOMATION"
3. Wait 3 minutes
4. Done! All mails sent automatically
```

### Scenario 2: Check Progress
```
1. While running, watch the:
   - Progress bar (0% to 100%)
   - Live console output
   - Current step description
2. Do NOT close the application
3. Wait until 100% and "✓ COMPLETED" message
```

### Scenario 3: View Results
```
1. After automation completes:
2. Click "📁 OPEN OUTPUT" to see archived files
3. Click "📋 VIEW LOGS" to see:
   - RUN_STATUS_*.xlsx (summary)
   - automation_*.log (details)
```

### Scenario 4: Handle Failures
```
1. Application automatically retries failed mails 3 times
2. If still fails, error is logged
3. Check status report for details
4. Fix the issue (wrong email, etc.)
5. Run automation again
```

---

## Troubleshooting

### Problem: "Master File: Not detected"

**Solution**:
1. Ensure `Input` folder exists
2. Place `Master.xlsx` inside `Input` folder
3. Restart application
4. It should auto-detect

### Problem: "Mail Map: Not found"

**Solution**:
1. Ensure `config` folder exists
2. Place `DSA_MAIL_MAP.xlsx` inside `config` folder
3. Restart application

### Problem: Mails not sending

**Solution**:
1. Ensure Microsoft Outlook is installed
2. Ensure Outlook is running
3. Check email addresses in DSA_MAIL_MAP.xlsx
4. Check internet connection
5. Try running as Administrator

### Problem: "No email found for DSA: ABC Finance"

**Solution**:
1. Open DSA_MAIL_MAP.xlsx
2. Check that DSA names match exactly:
   - Case sensitive
   - No extra spaces
   - Same spelling as Master.xlsx
3. Save and try again

### Problem: Application crashes

**Solution**:
1. Check Logs/ folder for error details
2. Ensure all dependencies are installed
3. Restart computer and try again
4. Run as Administrator

---

## File Format Requirements

### Master.xlsx
```
Required Column: "DSA"

Example:
┌─────────────────┬─────────────┬──────────────┬──────────────┐
│ DSA             │ Field1      │ Field2       │ Field3       │
├─────────────────┼─────────────┼──────────────┼──────────────┤
│ ABC Finance     │ Value1      │ Value2       │ Value3       │
│ ABC Finance     │ Value1      │ Value2       │ Value3       │
│ XYZ Associates  │ Value1      │ Value2       │ Value3       │
│ QRS Brokers     │ Value1      │ Value2       │ Value3       │
└─────────────────┴─────────────┴──────────────┴──────────────┘
```

### DSA_MAIL_MAP.xlsx
```
Required Columns: "DSA", "Email"

Example:
┌─────────────────┬──────────────────────────────┐
│ DSA             │ Email                        │
├─────────────────┼──────────────────────────────┤
│ ABC Finance     │ abc.finance@company.com      │
│ XYZ Associates  │ xyz.associates@company.com   │
│ QRS Brokers     │ qrs.brokers@company.com      │
│ PQR Consultants │ pqr.consultants@company.com  │
│ STU Partners    │ stu.partners@company.com     │
└─────────────────┴──────────────────────────────┘
```

---

## Tips for Best Results

1. **Prepare Files Before Running**
   - Ensure Master.xlsx is clean and error-free
   - Verify all DSA names match between files
   - Double-check email addresses

2. **Run During Off-Peak Hours**
   - Avoid sending many mails during work hours
   - Risk of network congestion is lower

3. **Keep Application Running**
   - Do not close during automation
   - Do not disconnect internet
   - Do not put computer to sleep

4. **Monitor Progress**
   - Watch the live console
   - Note any warnings or errors
   - Keep logs for records

5. **Archive Completed Runs**
   - Backup Logs folder periodically
   - Backup Archive folder for records
   - Maintain version history

---

## Need Help?

1. **Check Logs**
   - Click "📋 VIEW LOGS"
   - Look for error messages
   - Search for "ERROR" or "FAILED"

2. **Review Console Output**
   - Watch the live console during execution
   - Note exact error messages
   - Screenshot if needed

3. **Contact Support**
   - Provide error message from logs
   - Provide console output screenshot
   - Provide file samples (without sensitive data)

---

**Version**: 1.0.0
**Last Updated**: 2024
**Support**: BHFL Support Team
