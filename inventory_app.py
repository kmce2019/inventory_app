import tkinter
import customtkinter
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime

# ===================================================================================
# --- CONFIGURATION - PLEASE REVIEW CAREFULLY ---
# ===================================================================================

# 1. Define keywords to identify asset types from your filenames.
COMPUTER_KEYWORDS = ['computer', 'laptop', 'pc']
DESK_PHONE_KEYWORDS = ['desk', 'deskphone', 'voip']
CELL_PHONE_KEYWORDS = ['cell', 'mobile', 'iphone', 'android']

# 2. Define the column name in your asset files that contains the user's name.
USER_COLUMN_IN_ASSETS = 'Assigned User'

# 3. Map the final report columns to the actual column names in your source Excel files.
ASSET_FIELD_MAPPING = {
    # Final Report Column -> [Your Source Column Names]
    'Computer_Name':      ['Name'],
    'Computer_Model':     ['Model', 'Make', 'Device Class'],
    'Computer_Serial':    ['Serial Number'],
    'Desk_Phone_Ext':     ['extensionNumber'],
    'Cell_Phone_Model':   ['Model', 'Device Name'],
    'Cell_Phone_Number':  ['Cell Phone Number', 'Mobile Number']
}

# 4. Define users to EXCLUDE from the Inactive User report.
EXCLUDE_EXACT_USERS = [
    '101 pager', '101 ticketoffice', '102 loud ringer', '102 pager', '102 ticket office', '103 laporte ringer ata', '103 ticketoffice', 
    '104 loud ringer', '104 open user 1', '104 open user 2', '104 pager', '104 ticket room', '105 conferenceroom', '105 mainoffice', 
    '105 ticketroom', '106 driver\'s room', '106 foodgrade ticket office', '106 main office', '107-depot', '108 corrigo ticket room', '108 loudringer', 
    '108 pager', '108 ticket', '110 columbus ringer ata', '110 fg', '111 cincinnatiloudspeaker', '112 louisville ringer ata', 
    '112 ticket office', '113 loud ringer', '113 pager', '113 ticketroom', '114 pager', '117 sarnia user', '118 open user', '119 break room', 
    '119 loudringer', '119 pager', '119 ticketroom', '130 loud ringer', '130 pager', '131 main office', '131 loud ringer', '131 pager', '131 ticket', 
    '146 main', '147 main', '149 user1', '150 user3', '156 frontdesk', '165 user1', '165 wash bay 1', '190 main office', '195 line 1', 
    '203 ticket', '204 ticket office', '215 user 2', '220 dallas open user', '222 servicedesk', '222 ticketoffice', '223 csc', '223 ticket', 
    '223 ticket 2', '233 tankwash', '246 - second office', '246 - ticket office', '249 linda dunbar', '249 loudringer', '249 open user', '249 pager', 
    '254-foodgrade', '254 pager', '255 loud ringer', '255 pager', '256 loud ringer', '256 pager', '256 user', '267 loud ringer', '267 pager', 
    '267 ticketoffice', '269 ticket office', '275 houston open user 1', '275 loudringer', '275 pager', '276 ticket office', '279 loud pager', 
    '279 loud ringer', '282 ticket office', '290 ticketoffice', '292 ticket', '401.basfibc', '401-label', '401cmisauto', '401cmisship', '401 guests', 
    '401 page', '403-label', '403 conference room', '405 front office', '405 inventory', '4825670fffa4', '4825671bab3e', '482567bae484', 
    '482567bb3a56', '482567bb3fa4', '482567bb91a9', '482567bc014e', '503 manager', '506 paging', '508 sales office', '512 conference', 
    '512 conference 2', '512 paging', '514 conference', '520 paging', '521 paging', '526 paging', '526 upstairs', '527-conference-exec', 
    '527-conference-main', '527-washbay-1', '527-washbay-2', '527 paging', '527 upstairs', '528 paging', '528 tankcon', '529 paging', '530 biller', 
    '532 paging', '532 petroleum', '534 front office', '535 paging', '536 paging', '537 front office', '537 paging', '537 reception', 
    '537 service', '537 spare', '538 conference', '538 paging', '539 conference', '539 front office', '540 paging', '541 paging', 
    '541 service manager', '541 tank wash', '543 paging', '544 paging', '544 service writer', '547 paging', '550 paging', '555 paging', 
    '559 paging', '561 dan', '561 paging', '565 pager', '566 paging', '567 breakroom', '572 conference', '572 paging', '573 office', 
    '573 operator', '573 paging', '574 paging', '574 service', '586 conference', '586 paging', '586 sales office', '587 break room', 
    '587 paging', '589 paging', '596 office', '599 paging', '6161 algo', '665 user 1', '665 user 2', '701 tank wash', '701 truck drivers', 
    '702 labarea', '704 box gate', '704 conference main', '704 conference ops', '704 connex', '704 entrance gate', 'admin', 'administrator', 'ballast', 
    'camera', 'cameras', 'dieseltech', 'drive kiosk', 'euautomatedreports', 'frontlinescan', 'houston275', 'labelmaker', 'laporteticket', 
    'mailattachment', 'nablesa', 'nlr506', 'pdeangelis', 'plcuser', 'qhl-gibraltar', 'qrsreports', 'qualatv', 'sacramento534', 'stackcom', 
    'store248', 'supervisor', 'surface', 'tablet', 'tank251', 'toolbox', 'user', 'userx', 'wash computer 2'
]

EXCLUDE_KEYWORD_USERS = [
    'admin', 'shop', 'part', 'timeclock', 'time', 'tc', 'backparts', 
    'training', 'tickets', 'counter', 'diag', 'warehouse', 'joliet', 
    'karmak', 'kiosk', 'shipping', 'boasso'
]

# 5. Phone Usage Report Column Names
PHONE_USAGE_COLUMNS = {
    'user_id': 'Assigned User',
    'inbound_duration': 'Inbound Duration (seconds)',
    'outbound_duration': 'Outbound Duration (seconds)'
}
# ===================================================================================

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Advanced Inventory Exporter v4.3")
        self.geometry("800x700")

        # --- Data Storage ---
        self.processed_df = None
        self.active_pids = None
        self.name_to_id_map = None # To store the name mapping
        
        # --- Exclusion Lists are now part of the App class ---
        self.EXCLUDE_EXACT_USERS = EXCLUDE_EXACT_USERS
        self.EXCLUDE_KEYWORD_USERS = EXCLUDE_KEYWORD_USERS
        
        # --- Configure grid layout ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # --- Top Frame for File Loading ---
        self.load_frame = customtkinter.CTkFrame(self)
        self.load_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        self.load_frame.grid_columnconfigure(0, weight=1)

        self.master_button = customtkinter.CTkButton(self.load_frame, text="1. Upload Master Employee List", command=self.upload_master_file)
        self.master_button.grid(row=0, column=0, padx=10, pady=(10,5), sticky="ew")

        self.assets_button = customtkinter.CTkButton(self.load_frame, text="2. Upload Asset Files", command=self.upload_asset_files)
        self.assets_button.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        
        self.phone_usage_button = customtkinter.CTkButton(self.load_frame, text="3. Upload Phone Usage Report", command=self.upload_phone_usage_file)
        self.phone_usage_button.grid(row=2, column=0, padx=10, pady=5, sticky="ew")

        self.process_button = customtkinter.CTkButton(self.load_frame, text="4. Load & Process Files", command=self.load_and_process_files, fg_color="#108a36", hover_color="#0a5c24")
        self.process_button.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

        # --- Main Frame for Report Buttons ---
        self.report_frame = customtkinter.CTkFrame(self)
        self.report_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.report_frame.grid_columnconfigure((0, 1), weight=1)
        
        report_label = customtkinter.CTkLabel(self.report_frame, text="Generate & Export Reports (.csv)", font=customtkinter.CTkFont(weight="bold"))
        report_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

        self.report_buttons = []
        reports = {
            "Active Employee Inventory": self.generate_report_1,
            "Inactive User Inventory": self.generate_report_2,
            "Active - Computers Only": self.generate_report_3,
            "Active - Desk Phones Only": self.generate_report_4,
            "Active - Cell Phones Only": self.generate_report_5,
            "Active - Desk & Cell Phone Users": self.generate_report_6
        }
        row, col = 1, 0
        for text, command in reports.items():
            button = customtkinter.CTkButton(self.report_frame, text=text, command=command, state="disabled")
            button.grid(row=row, column=col, padx=10, pady=10, sticky="ew")
            self.report_buttons.append(button)
            col += 1
            if col > 1:
                col = 0
                row += 1
        
        self.phone_usage_report_button = customtkinter.CTkButton(self.report_frame, text="Desk Phone Usage Report", command=self.generate_phone_usage_report, state="disabled", fg_color="#c95100", hover_color="#8f3a00")
        self.phone_usage_report_button.grid(row=row+1, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        self.report_buttons.append(self.phone_usage_report_button)

        # --- Status Textbox ---
        self.textbox = customtkinter.CTkTextbox(self, height=100)
        self.textbox.grid(row=2, column=0, padx=20, pady=20, sticky="nsew")
        self.update_status("Welcome! Please load your files and then process them.")

    def update_status(self, message):
        timestamp = datetime.now().strftime("%I:%M:%S %p")
        self.textbox.insert("0.0", f"[{timestamp}] {message}\n")

    def upload_master_file(self):
        path = filedialog.askopenfilename(title="Select Master Spreadsheet", filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.master_file_path = path
            self.update_status(f"Master file selected: {path.split('/')[-1]}")

    def upload_asset_files(self):
        paths = filedialog.askopenfilenames(title="Select Asset Spreadsheets", filetypes=[("Excel", "*.xlsx *.xls")])
        if paths:
            self.asset_file_paths = list(paths)
            self.update_status(f"Selected {len(paths)} asset file(s).")
            
    def upload_phone_usage_file(self):
        path = filedialog.askopenfilename(title="Select Phone Usage Report", filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.phone_usage_file_path = path
            self.update_status(f"Phone usage report selected: {path.split('/')[-1]}")
    
    def load_and_process_files(self):
        if not hasattr(self, 'master_file_path') or not hasattr(self, 'asset_file_paths'):
            messagebox.showerror("Error", "Please select both a master file and at least one asset file.")
            return
        try:
            self.update_status("Processing... This may take a moment.")
            self.processed_df = None
            master_df = pd.read_excel(self.master_file_path)
            if 'Position ID' not in master_df.columns: raise ValueError("Master file must contain a 'Position ID' column.")
            name_cols = [col for col in master_df.columns if col != 'Position ID']
            self.name_to_id_map = {str(row[col]).lower().strip(): row['Position ID'] for index, row in master_df.iterrows() for col in name_cols if pd.notna(row[col])}
            self.active_pids = set(master_df['Position ID'].dropna())
            all_assets_list = []
            for f in self.asset_file_paths:
                asset_df = pd.read_excel(f)
                if USER_COLUMN_IN_ASSETS not in asset_df.columns: raise ValueError(f"File '{f.split('/')[-1]}' lacks '{USER_COLUMN_IN_ASSETS}' column.")
                asset_df['Position ID'] = asset_df[USER_COLUMN_IN_ASSETS].apply(lambda name: self.name_to_id_map.get(str(name).lower().strip()))
                asset_df['Asset Source'] = f.split('/')[-1]
                all_assets_list.append(asset_df)
            combined_assets_df = pd.concat(all_assets_list, ignore_index=True)
            self.processed_df = pd.merge(combined_assets_df, master_df, on='Position ID', how='left')
            self.processed_df[USER_COLUMN_IN_ASSETS] = self.processed_df[USER_COLUMN_IN_ASSETS].fillna("Unknown")
            for button in self.report_buttons: button.configure(state="normal")
            self.update_status("✅ Files processed successfully. Ready to generate reports.")
            messagebox.showinfo("Success", "Data loaded. You can now generate reports.")
        except Exception as e:
            self.update_status(f"❌ Error during processing: {e}")
            messagebox.showerror("Processing Error", f"An error occurred: {e}")

    def export_to_csv(self, df_to_export, report_name):
        if df_to_export.empty:
            self.update_status(f"No data found for report '{report_name}'. Nothing to export.")
            messagebox.showinfo("No Data", "The report you generated has no data.")
            return
        timestamp = datetime.now().strftime("%Y-%m-%d")
        filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")], initialfile=f"{report_name}_{timestamp}.csv", title="Save Report As")
        if filepath:
            df_to_export.to_csv(filepath, index=False)
            self.update_status(f"Successfully exported '{report_name}' to {filepath.split('/')[-1]}")
            messagebox.showinfo("Export Successful", f"Report saved to:\n{filepath}")

    def _build_wide_report(self, source_df, asset_types_to_include=None, merge_key='Position ID'):
        if source_df.empty:
            return pd.DataFrame()
        if asset_types_to_include is None:
            asset_types_to_include = ['Computer', 'Desk_Phone', 'Cell_Phone']
        cols_to_keep = [merge_key]
        if merge_key == 'Position ID':
            cols_to_keep.append(USER_COLUMN_IN_ASSETS)
        user_base_df = source_df[cols_to_keep].copy().drop_duplicates(subset=[merge_key])
        final_report_df = user_base_df
        asset_definitions = {
            'Computer': (COMPUTER_KEYWORDS, ['Computer_Name', 'Computer_Model', 'Computer_Serial']),
            'Desk_Phone': (DESK_PHONE_KEYWORDS, ['Desk_Phone_Ext']),
            'Cell_Phone': (CELL_PHONE_KEYWORDS, ['Cell_Phone_Model', 'Cell_Phone_Number'])
        }
        for asset_name, (keywords, fields) in asset_definitions.items():
            if asset_name not in asset_types_to_include:
                continue
            asset_mask = source_df['Asset Source'].str.lower().str.contains('|'.join(keywords), na=False)
            asset_df = source_df[asset_mask].copy()
            for field in fields:
                possible_cols = ASSET_FIELD_MAPPING.get(field, [])
                actual_col = next((col for col in possible_cols if col in asset_df.columns), None)
                if actual_col:
                    agg_data = asset_df.groupby(merge_key)[actual_col].agg(lambda x: ' | '.join(x.dropna().astype(str).unique())).rename(field)
                    final_report_df = pd.merge(final_report_df, agg_data, on=merge_key, how='left')
                else:
                    final_report_df[field] = None
        return final_report_df.fillna('')

    def generate_report_1(self):
        active_df = self.processed_df[self.processed_df['Position ID'].isin(self.active_pids)].copy()
        report_df = self._build_wide_report(active_df)
        self.export_to_csv(report_df, "Active_Employee_Inventory")

    def generate_report_2(self):
        inactive_df = self.processed_df[self.processed_df['Position ID'].isnull()].copy()
        inactive_df = inactive_df.dropna(subset=[USER_COLUMN_IN_ASSETS])
        inactive_df = inactive_df[~inactive_df[USER_COLUMN_IN_ASSETS].str.lower().isin(self.EXCLUDE_EXACT_USERS)]
        keyword_regex = '|'.join(self.EXCLUDE_KEYWORD_USERS)
        inactive_df = inactive_df[~inactive_df[USER_COLUMN_IN_ASSETS].str.lower().str.contains(keyword_regex, na=False)]
        report_df = self._build_wide_report(inactive_df, merge_key=USER_COLUMN_IN_ASSETS)
        self.export_to_csv(report_df, "Inactive_User_Inventory")

    def generate_report_3(self):
        mask = self.processed_df['Asset Source'].str.lower().str.contains('|'.join(COMPUTER_KEYWORDS), na=False)
        active_computers_df = self.processed_df[self.processed_df['Position ID'].isin(self.active_pids) & mask].copy()
        report_df = self._build_wide_report(active_computers_df, asset_types_to_include=['Computer'])
        self.export_to_csv(report_df, "Active_Computers_Only")
    
    def generate_report_4(self):
        mask = self.processed_df['Asset Source'].str.lower().str.contains('|'.join(DESK_PHONE_KEYWORDS), na=False)
        active_phones_df = self.processed_df[self.processed_df['Position ID'].isin(self.active_pids) & mask].copy()
        report_df = self._build_wide_report(active_phones_df, asset_types_to_include=['Desk_Phone'])
        self.export_to_csv(report_df, "Active_Desk_Phones_Only")

    def generate_report_5(self):
        mask = self.processed_df['Asset Source'].str.lower().str.contains('|'.join(CELL_PHONE_KEYWORDS), na=False)
        active_cells_df = self.processed_df[self.processed_df['Position ID'].isin(self.active_pids) & mask].copy()
        report_df = self._build_wide_report(active_cells_df, asset_types_to_include=['Cell_Phone'])
        self.export_to_csv(report_df, "Active_Cell_Phones_Only")

    def generate_report_6(self):
        active_df = self.processed_df[self.processed_df['Position ID'].isin(self.active_pids)].copy()
        has_desk = active_df['Asset Source'].str.lower().str.contains('|'.join(DESK_PHONE_KEYWORDS), na=False)
        has_cell = active_df['Asset Source'].str.lower().str.contains('|'.join(CELL_PHONE_KEYWORDS), na=False)
        desk_pids = set(active_df[has_desk]['Position ID'])
        cell_pids = set(active_df[has_cell]['Position ID'])
        pids_with_both = desk_pids.intersection(cell_pids)
        users_with_both_df = active_df[active_df['Position ID'].isin(pids_with_both)].copy()
        report_df = self._build_wide_report(users_with_both_df, asset_types_to_include=['Desk_Phone', 'Cell_Phone'])
        self.export_to_csv(report_df, "Active_Users_With_Desk_and_Cell_Phones")
        
    def generate_phone_usage_report(self):
        if not hasattr(self, 'phone_usage_file_path'):
            messagebox.showerror("Error", "Please upload the Phone Usage Report first.")
            return
        if not hasattr(self, 'name_to_id_map'):
            messagebox.showerror("Error", "Please load and process the main inventory files first.")
            return

        try:
            self.update_status("Generating phone usage report...")
            cfg = PHONE_USAGE_COLUMNS
            phone_df = pd.read_excel(self.phone_usage_file_path)

            # Define the exact columns needed from the phone report
            required_cols = [cfg['user_id'], cfg['inbound_duration'], cfg['outbound_duration']]

            # Check if all required columns exist
            for col in required_cols:
                if col not in phone_df.columns:
                    raise ValueError(f"Phone Usage Report is missing required column: '{col}'")

            # Get the set of all known active employee names and aliases (lowercase)
            active_names = {name for name, pid in self.name_to_id_map.items() if pid in self.active_pids}

            # Filter the phone usage report to include only active users
            report_df = phone_df[phone_df[cfg['user_id']].str.lower().isin(active_names)].copy()

            # Select only the columns requested for the final output
            final_report = report_df[required_cols]

            self.export_to_csv(final_report, "Desk_Phone_Usage_Duration")

        except Exception as e:
            self.update_status(f"❌ Error generating phone usage report: {e}")
            messagebox.showerror("Report Error", f"Could not generate phone usage report. Please check file format and configuration.\n\nError: {e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
