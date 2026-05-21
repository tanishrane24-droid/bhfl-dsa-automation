"""
BHFL DSA AUTOMATION - Main Application
One-click desktop automation for DSA mail distribution
DEMO MODE ENABLED - Sample data ready for manager presentation
"""

import sys
import os
import json
import threading
import traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from queue import Queue

# Core imports
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except:
    OUTLOOK_AVAILABLE = False

from pathlib import Path
import shutil
import logging

# Setup paths
BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "Input"
OUTPUT_DIR = BASE_DIR / "Output"
ARCHIVE_DIR = BASE_DIR / "Archive"
LOG_DIR = BASE_DIR / "Logs"
CONFIG_DIR = BASE_DIR / "config"
DEMO_DIR = BASE_DIR / "Demo_Output"

# DEMO MODE FLAG
DEMO_MODE = True  # Set to False to send actual emails

# Create directories
for dir_path in [INPUT_DIR, OUTPUT_DIR, ARCHIVE_DIR, LOG_DIR, CONFIG_DIR, DEMO_DIR]:
    dir_path.mkdir(exist_ok=True)

# Setup logging
log_file = LOG_DIR / f"automation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class DSAAutomationApp:
    """Main application class"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("BHFL DSA AUTOMATION")
        self.root.geometry("900x750")
        self.root.resizable(False, False)
        
        # Set icon if available
        try:
            icon_path = BASE_DIR / "assets" / "icon.ico"
            if icon_path.exists():
                self.root.iconbitmap(icon_path)
        except:
            pass
        
        # Style configuration
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'), foreground='#1F4788')
        style.configure('Subtitle.TLabel', font=('Arial', 11), foreground='#555')
        style.configure('Status.TLabel', font=('Courier', 9), foreground='#333')
        style.configure('DemoMode.TLabel', font=('Arial', 9, 'bold'), foreground='#FF6B35')
        style.configure('Run.TButton', font=('Arial', 11, 'bold'), padding=15)
        style.configure('Small.TButton', font=('Arial', 9))
        
        # Variables
        self.master_file = None
        self.dsa_mail_map = None
        self.is_running = False
        self.log_queue = Queue()
        
        self.setup_ui()
        self.auto_detect_files()
        self.start_log_monitor()
    
    def setup_ui(self):
        """Setup user interface"""
        # Header
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill='x', padx=20, pady=15)
        
        ttk.Label(header_frame, text="BHFL DSA AUTOMATION", style='Title.TLabel').pack(anchor='w')
        ttk.Label(header_frame, text="One-click DSA Mail Distribution System", style='Subtitle.TLabel').pack(anchor='w')
        
        if DEMO_MODE:
            ttk.Label(header_frame, text="🔵 DEMO MODE - Sample Data Ready", style='DemoMode.TLabel').pack(anchor='w', pady=5)
        
        # Status frame
        status_frame = ttk.LabelFrame(self.root, text="Status", padding=10)
        status_frame.pack(fill='x', padx=20, pady=10)
        
        self.master_label = ttk.Label(status_frame, text="Master File: Not detected", style='Status.TLabel')
        self.master_label.pack(anchor='w', pady=5)
        
        self.mail_map_label = ttk.Label(status_frame, text="Mail Map: Not detected", style='Status.TLabel')
        self.mail_map_label.pack(anchor='w', pady=5)
        
        if DEMO_MODE:
            demo_label = ttk.Label(status_frame, text="Mode: DEMO (No actual emails sent)", style='DemoMode.TLabel')
            demo_label.pack(anchor='w', pady=5)
        
        # Progress frame
        progress_frame = ttk.LabelFrame(self.root, text="Progress", padding=10)
        progress_frame.pack(fill='x', padx=20, pady=10)
        
        self.progress = ttk.Progressbar(progress_frame, mode='determinate', length=400)
        self.progress.pack(fill='x', pady=10)
        
        self.progress_label = ttk.Label(progress_frame, text="Ready", style='Status.TLabel')
        self.progress_label.pack(anchor='w', pady=5)
        
        # Console output
        console_frame = ttk.LabelFrame(self.root, text="Live Console Output", padding=10)
        console_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        self.console = scrolledtext.ScrolledText(
            console_frame,
            height=12,
            width=100,
            font=('Courier', 8),
            bg='#f0f0f0',
            fg='#000'
        )
        self.console.pack(fill='both', expand=True)
        
        # Buttons frame
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill='x', padx=20, pady=20)
        
        self.run_button = ttk.Button(
            button_frame,
            text="▶ RUN AUTOMATION",
            command=self.run_automation,
            style='Run.TButton'
        )
        self.run_button.pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="📁 OPEN OUTPUT", command=self.open_output, style='Small.TButton').pack(side='left', padx=5)
        ttk.Button(button_frame, text="📋 VIEW LOGS", command=self.view_logs, style='Small.TButton').pack(side='left', padx=5)
        ttk.Button(button_frame, text="❌ EXIT", command=self.root.quit, style='Small.TButton').pack(side='left', padx=5)
    
    def auto_detect_files(self):
        """Auto-detect Master.xlsx and DSA_MAIL_MAP.xlsx"""
        try:
            # Check for Master.xlsx
            master_path = INPUT_DIR / "Master.xlsx"
            if master_path.exists():
                self.master_file = master_path
                self.master_label.config(text=f"✓ Master File: {master_path.name}")
                self.log_console(f"✓ Master file detected: {master_path}")
            else:
                self.master_label.config(text="✗ Master File: Not found in Input folder")
                self.log_console("✗ Master.xlsx not found in Input folder")
                self.log_console("  Tip: Run 'python generate_demo_data.py' to create sample files")
            
            # Check for mail map
            mail_map_path = CONFIG_DIR / "DSA_MAIL_MAP.xlsx"
            if mail_map_path.exists():
                self.dsa_mail_map = mail_map_path
                self.mail_map_label.config(text=f"✓ Mail Map: {mail_map_path.name}")
                self.log_console(f"✓ Mail map detected: {mail_map_path}")
            else:
                self.mail_map_label.config(text="✗ Mail Map: Not found in config folder")
                self.log_console("✗ DSA_MAIL_MAP.xlsx not found in config folder")
                self.log_console("  Tip: Run 'python generate_demo_data.py' to create sample files")
        
        except Exception as e:
            self.log_console(f"✗ Auto-detection error: {str(e)}")
            logger.error(f"Auto-detection error: {traceback.format_exc()}")
    
    def run_automation(self):
        """Run the entire automation process"""
        if self.is_running:
            messagebox.showwarning("Warning", "Automation is already running!")
            return
        
        if not self.master_file or not self.dsa_mail_map:
            messagebox.showerror("Error", "Required files not detected!\n\n" +
                               "Please ensure:\n" +
                               "1. Input/Master.xlsx exists\n" +
                               "2. config/DSA_MAIL_MAP.xlsx exists\n\n" +
                               "To generate sample files, run:\n" +
                               "python generate_demo_data.py")
            return
        
        self.is_running = True
        self.run_button.config(state='disabled')
        self.console.delete('1.0', tk.END)
        self.progress['value'] = 0
        
        # Run in separate thread
        thread = threading.Thread(target=self.automation_workflow, daemon=True)
        thread.start()
    
    def automation_workflow(self):
        """Complete automation workflow"""
        try:
            stats = {
                'total_dsa': 0,
                'mails_sent': 0,
                'failed': 0,
                'errors': []
            }
            
            self.log_console("=" * 60)
            if DEMO_MODE:
                self.log_console("STARTING AUTOMATION WORKFLOW (DEMO MODE)")
            else:
                self.log_console("STARTING AUTOMATION WORKFLOW")
            self.log_console("=" * 60)
            
            # Step 1: Read Master file
            self.progress_label.config(text="Reading Master file...")
            self.progress['value'] = 10
            self.log_console("\n[1/6] Reading Master.xlsx...")
            
            master_df = pd.read_excel(self.master_file)
            self.log_console(f"✓ Read {len(master_df)} records from Master.xlsx")
            
            # Step 2: Identify unique DSAs
            self.progress_label.config(text="Detecting DSA names...")
            self.progress['value'] = 20
            self.log_console("\n[2/6] Identifying unique DSA names...")
            
            dsa_names = master_df['DSA Name'].unique()
            stats['total_dsa'] = len(dsa_names)
            self.log_console(f"✓ Detected {stats['total_dsa']} unique DSAs:")
            for dsa in dsa_names:
                self.log_console(f"  - {dsa}")
            
            # Step 3: Create DSA-wise files
            self.progress_label.config(text="Creating DSA-wise files...")
            self.progress['value'] = 30
            self.log_console("\n[3/6] Creating DSA-wise Excel files...")
            
            dsa_files = {}
            for dsa in dsa_names:
                dsa_df = master_df[master_df['DSA Name'] == dsa]
                file_name = f"{dsa.replace(' ', '_')}.xlsx"
                file_path = OUTPUT_DIR / file_name
                
                dsa_df.to_excel(file_path, index=False, sheet_name='Data')
                dsa_files[dsa] = file_path
                self.log_console(f"✓ Created {file_name} ({len(dsa_df)} records)")
            
            # Step 4: Load mail map
            self.progress_label.config(text="Loading mail configuration...")
            self.progress['value'] = 40
            self.log_console("\n[4/6] Loading DSA mail map...")
            
            mail_map_df = pd.read_excel(self.dsa_mail_map)
            mail_map = dict(zip(mail_map_df['DSA Name'], mail_map_df['Mail ID']))
            self.log_console(f"✓ Loaded mail map for {len(mail_map)} DSAs")
            
            # Step 5: Create and send Outlook mails
            self.progress_label.config(text="Sending mails via Outlook...")
            self.progress['value'] = 50
            
            if DEMO_MODE:
                self.log_console("\n[5/6] DEMO MODE: Preparing mail preview files (NOT sending)...")
            else:
                self.log_console("\n[5/6] Sending Outlook mails with attachments...")
            
            status_records = []
            
            if not DEMO_MODE and OUTLOOK_AVAILABLE:
                outlook = win32com.client.Dispatch('Outlook.Application')
            else:
                outlook = None
            
            for dsa, file_path in dsa_files.items():
                mail_sent = False
                retry_count = 0
                max_retries = 3
                
                while retry_count < max_retries and not mail_sent:
                    try:
                        if dsa not in mail_map:
                            raise ValueError(f"No email found for DSA: {dsa}")
                        
                        recipient = mail_map[dsa]
                        
                        if DEMO_MODE:
                            # DEMO MODE: Create sample mail file instead of sending
                            dsa_clean = dsa.lower().replace(' ', '_')
                            mail_preview = f"""
================================================================================
MAIL PREVIEW - DEMO MODE
================================================================================

TO: {recipient}
SUBJECT: DSA Data Distribution - {dsa}
DATE: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

================================================================================

Dear {dsa},

Please find attached the data distribution file for your records.

File: {file_path.name}
Records: {len(master_df[master_df['DSA Name'] == dsa])}
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Best regards,
BHFL DSA Automation System

================================================================================
NOTE: This is a DEMO MODE email. No actual mail was sent.
================================================================================
                            """
                            
                            demo_file = DEMO_DIR / f"Mail_{dsa_clean}.txt"
                            with open(demo_file, 'w') as f:
                                f.write(mail_preview)
                            
                            stats['mails_sent'] += 1
                            mail_sent = True
                            self.log_console(f"✓ Mail preview created for {dsa} ({recipient})")
                            
                            status_records.append({
                                'Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                'DSA': dsa,
                                'Mail': recipient,
                                'Status': 'SUCCESS (DEMO)',
                                'Remarks': 'Mail preview created (Demo Mode - not sent)'
                            })
                        
                        else:
                            # PRODUCTION MODE: Send actual email
                            if not outlook:
                                raise Exception("Outlook not available. Please install Microsoft Outlook.")
                            
                            mail = outlook.CreateItem(0)
                            mail.To = recipient
                            mail.Subject = f"DSA Data Distribution - {dsa}"
                            mail.Body = f"""Dear {dsa},

Please find attached the data distribution file for your records.

File: {file_path.name}
Records: {len(master_df[master_df['DSA Name'] == dsa])}

Best regards,
BHFL DSA Automation System"""
                            
                            mail.Attachments.Add(str(file_path))
                            mail.Send()
                            
                            stats['mails_sent'] += 1
                            mail_sent = True
                            self.log_console(f"✓ Mail sent to {dsa} ({recipient})")
                            
                            status_records.append({
                                'Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                'DSA': dsa,
                                'Mail': recipient,
                                'Status': 'SUCCESS',
                                'Remarks': 'Mail sent successfully'
                            })
                    
                    except Exception as e:
                        retry_count += 1
                        if retry_count < max_retries:
                            self.log_console(f"⚠ Retry {retry_count}/{max_retries} for {dsa}...")
                        else:
                            stats['failed'] += 1
                            error_msg = str(e)
                            stats['errors'].append(f"{dsa}: {error_msg}")
                            self.log_console(f"✗ Failed for {dsa}: {error_msg}")
                            
                            status_records.append({
                                'Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                'DSA': dsa,
                                'Mail': mail_map.get(dsa, 'N/A'),
                                'Status': 'FAILED',
                                'Remarks': error_msg[:50]
                            })
            
            # Step 6: Archive files
            self.progress_label.config(text="Archiving processed files...")
            self.progress['value'] = 75
            self.log_console("\n[6/6] Archiving processed files...")
            
            for file_path in dsa_files.values():
                if file_path.exists():
                    shutil.move(str(file_path), str(ARCHIVE_DIR / file_path.name))
                    self.log_console(f"✓ Archived {file_path.name}")
            
            # Create status file
            self.progress['value'] = 85
            self.log_console("\nGenerating status report...")
            
            status_df = pd.DataFrame(status_records)
            status_file = LOG_DIR / f"RUN_STATUS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            status_df.to_excel(status_file, index=False, sheet_name='Status')
            self.log_console(f"✓ Status file created: {status_file.name}")
            
            # Final dashboard
            self.progress['value'] = 100
            self.progress_label.config(text="✓ COMPLETED")
            
            self.log_console("\n" + "=" * 60)
            if DEMO_MODE:
                self.log_console("AUTOMATION COMPLETED (DEMO MODE)")
            else:
                self.log_console("AUTOMATION COMPLETED SUCCESSFULLY")
            self.log_console("=" * 60)
            self.log_console(f"\nSUMMARY DASHBOARD:")
            self.log_console(f"  Total DSAs:      {stats['total_dsa']}")
            self.log_console(f"  Mails Processed: {stats['mails_sent']}")
            self.log_console(f"  Failed:          {stats['failed']}")
            self.log_console(f"  Success Rate:    {(stats['mails_sent']/stats['total_dsa']*100):.1f}%")
            
            if DEMO_MODE:
                self.log_console(f"\n  📁 Demo mail previews saved to: Demo_Output/")
            
            if stats['errors']:
                self.log_console(f"\nErrors encountered:")
                for error in stats['errors']:
                    self.log_console(f"  - {error}")
            
            self.log_console("\n" + "=" * 60)
            
            # Show completion popup
            mode_text = "DEMO MODE" if DEMO_MODE else ""
            self.root.after(0, lambda: messagebox.showinfo(
                "SUCCESS",
                f"AUTOMATION COMPLETED {mode_text}\n\n" +
                f"Total DSAs: {stats['total_dsa']}\n" +
                f"Processed: {stats['mails_sent']}\n" +
                f"Failed: {stats['failed']}\n" +
                f"Success Rate: {(stats['mails_sent']/stats['total_dsa']*100):.1f}%"
            ))
        
        except Exception as e:
            error_msg = f"AUTOMATION FAILED: {str(e)}\n\n{traceback.format_exc()}"
            self.log_console(f"\n✗ {error_msg}")
            logger.error(error_msg)
            self.root.after(0, lambda: messagebox.showerror("ERROR", f"Automation failed:\n{str(e)}"))
        
        finally:
            self.is_running = False
            self.run_button.config(state='normal')
    
    def log_console(self, message):
        """Log message to console"""
        self.console.insert(tk.END, message + "\n")
        self.console.see(tk.END)
        self.root.update()
        logger.info(message)
    
    def open_output(self):
        """Open output folder"""
        try:
            os.startfile(OUTPUT_DIR)
        except:
            try:
                os.system(f'open "{OUTPUT_DIR}"')  # For Mac
            except:
                messagebox.showerror("Error", f"Could not open: {OUTPUT_DIR}")
    
    def view_logs(self):
        """View logs folder"""
        try:
            os.startfile(LOG_DIR)
        except:
            try:
                os.system(f'open "{LOG_DIR}"')  # For Mac
            except:
                messagebox.showerror("Error", f"Could not open: {LOG_DIR}")
    
    def start_log_monitor(self):
        """Monitor log queue"""
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_console(msg)
        except:
            pass
        self.root.after(100, self.start_log_monitor)


def main():
    """Main entry point"""
    root = tk.Tk()
    app = DSAAutomationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
