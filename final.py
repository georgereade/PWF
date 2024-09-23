import requests
from requests.auth import HTTPBasicAuth
import os
from dotenv import load_dotenv
import pandas as pd

from datetime import datetime

import openpyxl
from openpyxl.utils import get_column_letter

load_dotenv()

# Define your API credentials
API_KEY = os.getenv('PWF_API_KEY')  # Replace with your actual ProWorkflow API key
BASE_URL = 'https://api.proworkflow.net'  # Base URL for the ProWorkflow API
USERNAME = os.getenv('PWF_USERNAME')  # Replace with your ProWorkflow username
PASSWORD = os.getenv('PWF_PASSWORD')  # Replace with your ProWorkflow password

# Headers for authentication
headers = {
    'Content-Type': 'application/json',
    'apikey': API_KEY,
}

pd.set_option('display.max_colwidth', None)

trackedfrom = '2024-08-01'  # Specify start date for the time range
trackedto = '2024-08-31'    # Specify end date for the time range

# Function to get all contacts of type 'staff'
def get_staff_contacts():
    url = f'{BASE_URL}/contacts'
    response = requests.get(url, headers=headers, auth=HTTPBasicAuth(USERNAME, PASSWORD))
    if response.status_code == 200:
        contacts = response.json()['contacts']
        print("Collected staff names...")
        return [contact for contact in contacts if contact['type'] == 'staff']
    else:
        raise Exception(f"Error fetching contacts: {response.status_code}, {response.text}")

# First request: Function to get time tracked by a contact for a specific period
def get_contact_time_totals(contact_id, trackedfrom, trackedto):
    url = f'{BASE_URL}/contacts/{contact_id}/time?trackedfrom={trackedfrom}&trackedto={trackedto}&subtotals=project'
    response = requests.get(url, headers=headers, auth=HTTPBasicAuth(USERNAME, PASSWORD))
    if response.status_code == 200:
        return response.json()['subtotals']
    else:
        raise Exception(f"Error fetching time for contact {contact_id}: {response.status_code}, {response.text}")

# Second request: Function to get task-specific time details by contact
def get_contact_task_details(contact_id, trackedfrom, trackedto):
    url = f'{BASE_URL}/contacts/{contact_id}/time?trackedfrom={trackedfrom}&trackedto={trackedto}&fields=dates,project,task,notes,contact'
    response = requests.get(url, headers=headers, auth=HTTPBasicAuth(USERNAME, PASSWORD))
    if response.status_code == 200:
        return response.json()['timerecords']
    else:
        raise Exception(f"Error fetching task details for contact {contact_id}: {response.status_code}, {response.text}")

# Function to calculate time spent between start and end times
def calculate_time_spent(start_time, end_time):
    start = pd.to_datetime(start_time)
    end = pd.to_datetime(end_time)
    time_spent = (end - start).total_seconds() / 60  # Time spent in minutes
    return f"{int(time_spent // 60)}:{int(time_spent % 60):02d}"  # Return formatted HH:MM

def format_time(iso_time):
    return pd.to_datetime(iso_time).strftime('%H:%M')

# Function to get project details
def get_project_details(project_id):
    url = f'{BASE_URL}/projects/{project_id}'
    response = requests.get(url, headers=headers, auth=HTTPBasicAuth(USERNAME, PASSWORD))
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f"Error fetching project {project_id}: {response.status_code}, {response.text}")

# Function to process both time totals and task details
def process_time_per_contact(trackedfrom, trackedto):
    staff_contacts = get_staff_contacts()
    project_data_totals = {}
    project_data_tasks = {}
    
    print("Calculating time per staff member...")

    for contact in staff_contacts:
        contact_id = contact['id']
        contact_name = f"{contact['firstname']} {contact['lastname']}"

        # First API call: Get time totals per project for the specific contact
        time_totals = get_contact_time_totals(contact_id, trackedfrom, trackedto)

        # Second API call: Get task details per project for the specific contact
        task_details = get_contact_task_details(contact_id, trackedfrom, trackedto)

        # Process time totals per project
        for record in time_totals:
            project_name = record['projecttitle']
            time_spent = record['timetracked']
            project_number = record['projectnumber']
            project_id = record['projectid']
            project_details = get_project_details(project_id)
            category_name = project_details['project']['categoryname']

            # Filter for categories 'On Hold' or 'Current Timed Projects'
            if category_name in ['On Hold', 'Current Timed Projects']:
                formatted_date = pd.to_datetime(trackedfrom).strftime('%b %Y')
                
                # Create a new entry for the project if it doesn't exist
                if project_name not in project_data_totals:
                    project_data_totals[project_name] = []
                
                project_data_totals[project_name].append({
                    'Person': contact_name,
                    'Project Number': project_number,
                    f'Total Time Spent ({formatted_date})': f"{time_spent // 60}:{time_spent % 60:02d}"
                })

        # Process task details per project
        for record in task_details:
            project_name = record['projecttitle']
            task_name = record['taskname']
            start_time = record['starttime']
            end_time = record['endtime']
            task_date = pd.to_datetime(start_time.split('T')[0]).strftime('%b %d, %Y')
            notes = record.get('notes', '')
            time_spent = calculate_time_spent(start_time, end_time)
            formatted_start_time = format_time(start_time)
            formatted_end_time = format_time(end_time)

            # Add task-specific details to the project
            if project_name not in project_data_tasks:
                project_data_tasks[project_name] = []
            
            project_data_tasks[project_name].append({
                'Task Date': task_date,
                'Staff': contact_name,
                'Task Name': task_name,
                'Time Record': notes,
                'Start': formatted_start_time,
                'Finish': formatted_end_time,
                'Time Spent': time_spent
            })

    return project_data_totals, project_data_tasks

# Main function to write separate Excel files per project
def main():
    # Process time for all staff contacts
    project_data_totals, project_data_tasks = process_time_per_contact(trackedfrom, trackedto)

    # Create output directory if it doesn't exist
    output_dir = 'output/projects'
    os.makedirs(output_dir, exist_ok=True)

    # Write each project's data to a separate Excel file
    for project_name, records in project_data_totals.items():
        df_totals = pd.DataFrame(records)
        df_tasks = pd.DataFrame(project_data_tasks.get(project_name, []))  # Get task details, if available
        
        # Replace invalid characters in project names that can't be used in file names
        safe_project_name = "".join([c if c.isalnum() or c in (' ', '-', '_') else '_' for c in project_name])
        formatted_date = pd.to_datetime(trackedfrom).strftime('%b %Y')
        project_number = records[0]['Project Number']
        excel_file = f'{output_dir}/{project_number} - {safe_project_name} {formatted_date} Timesheet.xlsx'

        # Write DataFrame to Excel
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df_totals.to_excel(writer, sheet_name='Totals', index=False)
            df_tasks.to_excel(writer, sheet_name='Task Details', index=False)

            # Adjust the width of each column to fit the content
            for sheet_name in ['Totals', 'Task Details']:
                worksheet = writer.sheets[sheet_name]
                df = df_totals if sheet_name == 'Totals' else df_tasks
                for col_idx, col in enumerate(df.columns, 1):
                    max_length = max(df[col].astype(str).map(len).max(), len(col))  # Find the max length in the column
                    worksheet.column_dimensions[get_column_letter(col_idx)].width = max_length + 2  # Add a little padding

        print(f"Time data for project '{project_name}' saved to {excel_file}")

if __name__ == "__main__":
    main()
