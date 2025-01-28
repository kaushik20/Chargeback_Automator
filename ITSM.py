import os
import logging
import pandas as pd
import win32com.client as win32
from datetime import datetime, timedelta
import re

logging.basicConfig(level=logging.INFO)

class WorkOrderReportProcessor:
   def __init__(self, input_file_path):
      self.input_file_path = input_file_path
      self.excluded_customer = "2122"
      self.mrc_values = {"Create the user id- Generic": '$39.15', "Microsoft Office E1 to E3 License Assignment - Task": '$12.94', "Power BI Pro License Assignment - Task": '$12.94', "Microsoft Project Premium License Assignment - Task": 17.39, "Assign License - Copilot": '$31.00'}
      self.default_save_dir = os.path.join(os.path.expanduser("~"), "Documents")
      self.df = None
      self.filtered_dataframes = []

   def load_data(self):
      try:
         self.df = pd.read_excel(self.input_file_path, sheet_name='WO Report')
         self.df['Closure Code'] = self.df['Closure Code'].str.strip()
         self.df['Customer'] = self.df['Customer'].str.strip()
         self.df['Actual Resolution Time'] = pd.to_datetime(self.df['Actual Resolution Time'], errors='coerce')
         logging.info(f"Data loaded: {self.df.shape[0]} rows")
      except Exception as e:
         logging.error(f"Error loading data: {e}")
         raise

   def filter_by_date(self):
      today = datetime.today()
      first_day_of_current_month = today.replace(day=1)
      last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
      first_day_of_previous_month = last_day_of_previous_month.replace(day=1)
      self.df = self.df[(self.df['Actual Resolution Time'] >= first_day_of_previous_month) & (self.df['Actual Resolution Time'] <= last_day_of_previous_month + timedelta(hours=23, minutes=59, seconds=59))]
      logging.info(f"Rows after date filter: {self.df.shape[0]} rows")

   def filter_by_closure_code(self):
      self.df = self.df[self.df['Closure Code'].str.strip() == 'Request fulfilled successfully']
      logging.info(f"Rows after 'Closure Code' filter: {self.df.shape[0]} rows")

   def filter_by_customer(self):
      self.df = self.df[self.df['Customer'].str.strip() != self.excluded_customer]
      logging.info(f"Rows after excluding Customer '2122': {self.df.shape[0]} rows")

   def process_chargeback(self, description_key, mrc_value_key):
    email_pattern = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
    filtered_rows = []
    for index, row in self.df.iterrows():
        if row['Description'] == description_key and pd.notna(row['Solution']):
            # Extract unique email addresses
            email_matches = set(email_pattern.findall(row['Solution']))
            email_count = len(email_matches)
            if email_count > 0:
                if mrc_value_key not in self.mrc_values:
                    logging.warning(f"MRC value for '{mrc_value_key}' not found. Skipping row {index}.")
                    continue
                mrc_per_user = float(self.mrc_values[mrc_value_key].replace('$', ''))
                total_mrc = mrc_per_user * email_count
                new_row = row.copy()
                new_row['MRC'] = f"${total_mrc:.2f}"
                new_row['Email Count'] = email_count  # Add email count for debugging
                filtered_rows.append(new_row)
    df_filtered = pd.DataFrame(filtered_rows)
    self.filtered_dataframes.append(df_filtered)
    logging.info(f"Processed '{description_key}' with {len(filtered_rows)} rows.")

   def filter_assign_license_copilot(self):
      self.process_chargeback('Assign License - Copilot', 'Assign License - Copilot')

   def filter_create_user_id_generic(self):
      self.process_chargeback('Create the user id- Generic', 'Create the user id- Generic')

   def filter_office_e1_to_e3_license(self):
      self.process_chargeback('Microsoft Office E1 to E3 License Assignment - Task', 'Microsoft Office E1 to E3 License Assignment - Task')

   def filter_power_bi_pro_license(self):
      self.process_chargeback('Power BI Pro License Assignment - Task', 'Power BI Pro License Assignment - Task')

   def filter_project_professional_license(self):
      self.process_chargeback('Microsoft Project Premium License Assignment - Task', 'Microsoft Project Premium License Assignment - Task')

   def combine_filtered_data(self):
      combined_df = pd.concat(self.filtered_dataframes, ignore_index=True)
      logging.info(f"Total combined rows: {combined_df.shape[0]}")
      self.df = combined_df

   def select_and_modify_columns(self):
      selected_columns = ['Customer', 'Service Request No.', 'Caller', 'Description', 'MRC']
      self.df = self.df[selected_columns]
      current_date = datetime.now()
      first_day_of_month = current_date.replace(day=1).strftime("%B-%Y")
      self.df['Start Time'] = first_day_of_month
      logging.info("Columns selected and modified successfully.")

   def save_to_excel(self):
      try:
         current_date = datetime.now()
         month = current_date.strftime("%B")
         output_file_path = os.path.join(self.default_save_dir, f"Chargeback_{month}_2024.xlsx")
         os.makedirs(self.default_save_dir, exist_ok=True)
         self.df.to_excel(output_file_path, index=False)
         self.output_file_path = output_file_path
         logging.info(f"Filtered data saved to {self.output_file_path}.")
      except Exception as e:
         logging.error(f"Error saving to Excel: {e}")
         raise

   def send_email(self):
      try:
         outlook = win32.Dispatch('outlook.application')
         mail = outlook.CreateItem(0)
         namespace = outlook.GetNamespace("MAPI")
         account_email = None
         for account in namespace.Accounts:
            account_email = account.SmtpAddress
            break  
         if not account_email:
            raise Exception("Unable to retrieve the current user's Outlook email address.")
         current_date = datetime.now()
         month = current_date.strftime("%B")
         mail.To = account_email
         mail.Subject = f'Chargeback Report for {month} 2024'
         mail.Body = 'Please find the attached work order report.'
         mail.Attachments.Add(self.output_file_path)
         mail.Send()
         logging.info("Email sent successfully to the current Outlook user.")
      except Exception as e:
         logging.error(f"Error sending email: {e}")
         raise

   def automate_ITSM(self):
      try:
         self.load_data()
         self.filter_by_date()
         self.filter_by_closure_code()
         self.filter_by_customer()  
         self.filter_assign_license_copilot()
         self.filter_create_user_id_generic()
         self.filter_office_e1_to_e3_license()
         self.filter_power_bi_pro_license()
         self.filter_project_professional_license()
         self.combine_filtered_data()
         self.select_and_modify_columns()
         self.save_to_excel()
         self.send_email()
      except Exception as e:
         logging.error(f"Automation process failed: {e}")

if __name__ == "__main__":
   capital = WorkOrderReportProcessor("d:\\Users\\kaushikc\\Downloads\\WO Report.xlsx")
   capital.automate_ITSM()
