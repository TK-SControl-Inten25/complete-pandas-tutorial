def process_navigator_sheet():
    try:
        # Read Navigator Excel file
        nav_file = file_mapping["adv_st"]["path"]
        
        # Read Excel file with openpyxl
        wb = openpyxl.load_workbook(nav_file)
        ws = wb['Navigator']
        
        # Create a list to store non-strikethrough row indices
        rows_to_keep = []
        
        # Check each row for strikethrough formatting
        for row_idx, row in enumerate(ws.rows):
            # Check first cell in each row
            cell = row[0]
            
            # If the font doesn't have strikethrough or cell is empty, keep the row
            if not cell.font.strike or cell.value is None:
                rows_to_keep.append(row_idx)
        
        # Read the Excel file into DataFrame, keeping only non-strikethrough rows
        nav_df = pd.read_excel(nav_file, sheet_name="Navigator", skiprows=lambda x: x not in rows_to_keep)
        
        event_idx = nav_df.columns.get_loc('Event')
        nav_df.insert(event_idx + 1, 'Config', False)

        alarm_idx = nav_df.columns.get_loc('Event log when alarm code(s) set')
        new_columns = [
            'Event Log When Parameter Changes?',
            'Auto Log at Power ON?',
            'Auto Log at Noon (12:05 PM)?'
        ]
        
        # Initialize new columns with 'FALSE'
        for idx, col in enumerate(reversed(new_columns)):
            nav_df.insert(alarm_idx + 1, col, 'FALSE')

        # Process Additional Reqs column
        for idx, row in nav_df.iterrows():
            additional_reqs = str(row['Additional Reqs']).lower()  # Convert to string and lowercase for better matching
            
            # Check for noon logging
            if any(keyword in additional_reqs for keyword in ['12:05', 'midday', 'noon']):
                nav_df.at[idx, 'Auto Log at Noon (12:05 PM)?'] = 'TRUE'
            
            # Check for power on/off logging
            if any(keyword in additional_reqs for keyword in ['power on', 'power off', 'power-on', 'power-off', 'poweron', 'poweroff']):
                nav_df.at[idx, 'Auto Log at Power ON?'] = 'TRUE'
            
            # Check for parameter changes logging
            if any(keyword in additional_reqs for keyword in ['parameter changes', 'param changes', 'value change', 'changes']):
                nav_df.at[idx, 'Event Log When Parameter Changes?'] = 'TRUE'

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(output_dir, f'Navigator_Analysis_{timestamp}.xlsx')

        nav_df.to_excel(output_file, index=False)
        print(f"Processing complete. Output saved to: {output_file}")
        print(f"Removed {ws.max_row - len(rows_to_keep)} strikethrough rows")

    except Exception as e:
        print(f"Error processing Navigator sheet: {str(e)}")

# Execute the function
process_navigator_sheet()

# -------------------------------------------------------------------- #

def process_deet_sheet():
    try:
        # Read DEET Excel file
        deet_file = file_mapping["DEET_ST"]["path"]
        
        # Read Excel file with openpyxl
        wb = openpyxl.load_workbook(deet_file)
        ws = wb['DEET']
        
        # Create a list to store non-strikethrough row indices
        rows_to_keep = []
        
        # Check each row for strikethrough formatting
        for row_idx, row in enumerate(ws.rows):
            cell = row[0]
            if not cell.font.strike or cell.value is None:
                rows_to_keep.append(row_idx)
        
        # Read the Excel file into DataFrame
        deet_df = pd.read_excel(deet_file, sheet_name="DEET", skiprows=lambda x: x not in rows_to_keep)
        
        # Filter rows based on Platform column
        deet_df = deet_df[deet_df['Platform'].str.contains('DEET|Navigator/DEET', case=False, na=False)]
        
        # Add columns after "Default Log"
        default_log_idx = deet_df.columns.get_loc('Default Log')
        new_columns = ['TIME', 'Event', 'Config']
        
        # Insert columns after Default Log with default FALSE values
        for idx, col in enumerate(reversed(new_columns)):
            deet_df.insert(default_log_idx + 1, col, 'FALSE')
        
        # Process Default Log column
        for idx, row in deet_df.iterrows():
            default_log = str(row['Default Log']).lower()  # Convert to string and lowercase for better matching
            
            # Check for Timed/Periodic logging
            if any(keyword in default_log for keyword in ['timed', 'periodic']):
                deet_df.at[idx, 'TIME'] = 'TRUE'
            
            # Check for Event logging
            if 'event' in default_log:
                deet_df.at[idx, 'Event'] = 'TRUE'
            
            # Check for Configuration logging
            if 'reefer configuration' in default_log:
                deet_df.at[idx, 'Config'] = 'TRUE'
        
        # Add columns after "Event log when alarm code(s) set"
        alarm_idx = deet_df.columns.get_loc('Event log when alarm code(s) set')
        log_columns = [
            'Event Log When Parameter Changes?',
            'Auto Log at Power ON?',
            'Auto Log at Noon (12:05 PM)?'
        ]
        
        # Insert columns after Event log when alarm code(s) set with default FALSE values
        for idx, col in enumerate(reversed(log_columns)):
            deet_df.insert(alarm_idx + 1, col, 'FALSE')

        # Process Additional Reqs column
        for idx, row in deet_df.iterrows():
            additional_reqs = str(row['Additional Reqs']).lower()  # Convert to string and lowercase for better matching
            
            # Check for noon/midday logging
            if any(keyword in additional_reqs for keyword in ['12:05', 'midday', 'noon', 'middle of day']):
                deet_df.at[idx, 'Auto Log at Noon (12:05 PM)?'] = 'TRUE'
            
            # Check for power on/startup logging
            if any(keyword in additional_reqs for keyword in ['power on', 'power-on', 'poweron', 'start up', 'startup', 'power cycle']):
                deet_df.at[idx, 'Auto Log at Power ON?'] = 'TRUE'
            
            # Check for parameter/value changes logging
            if any(keyword in additional_reqs for keyword in ['parameter', 'param', 'value change', 'state change', 'changes']):
                deet_df.at[idx, 'Event Log When Parameter Changes?'] = 'TRUE'

        # Remove Platform and Default Log columns
        deet_df = deet_df.drop(['Platform', 'Default Log'], axis=1)

        # Create output directory and save file
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(output_dir, f'DEET_Analysis_{timestamp}.xlsx')
        deet_df.to_excel(output_file, index=False)
        
        print(f"Processing complete. Output saved to: {output_file}")
        print(f"Removed {ws.max_row - len(deet_df)} rows (strikethrough + non-DEET platform)")

    except Exception as e:
        print(f"Error processing DEET sheet: {str(e)}")

# Execute the function
process_deet_sheet()