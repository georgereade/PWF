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

trackedfrom = '2023-01-01'  # Specify start date for the time range
trackedto = '2023-01-31'    # Specify end date for the time range

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

# Function to get time tracked by a contact for a specific period, broken down by tasks
def get_contact_time(contact_id, trackedfrom, trackedto):
    url = f'{BASE_URL}/contacts/{contact_id}/time?trackedfrom={trackedfrom}&trackedto={trackedto}&subtotals=project,task,day'
    response = requests.get(url, headers=headers, auth=HTTPBasicAuth(USERNAME, PASSWORD))
    if response.status_code == 200:
        return response.json()['subtotals']  # Assuming time data is under 'subtotals'
    else:
        raise Exception(f"Error fetching time for contact {contact_id}: {response.status_code}, {response.text}")

# Function to convert time from minutes to hours and minutes
def convert_time(minutes):
    hours = minutes // 60
    remaining_minutes = minutes % 60
    return f"{hours}:{remaining_minutes:02d}"

# Function to get project details
def get_project_details(project_id):
    url = f'{BASE_URL}/projects/{project_id}'
    response = requests.get(url, headers=headers, auth=HTTPBasicAuth(USERNAME, PASSWORD))
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f"Error fetching project {project_id}: {response.status_code}, {response.text}")

# Function to process time data for each staff member per project and task
def process_time_per_contact(trackedfrom, trackedto):
    staff_contacts = get_staff_contacts()
    project_data = {}
    print("Calculating time per staff member...")

    for contact in staff_contacts:
        contact_id = contact['id']
        contact_name = f"{contact['firstname']} {contact['lastname']}"

        # Get time tracked for the specific contact within the time range, broken down by project and task
        time_data = get_contact_time(contact_id, trackedfrom, trackedto)

        # Process time data per project and task
        for record in time_data:
            project_name = record['projecttitle']
            task_name = record['taskname']
            time_spent = record['timetracked']
            project_id = record['projectid']
            project_number = record['projectnumber']
            project_details = get_project_details(project_id)
            category_name = project_details['project']['categoryname']

            # Assuming the API provides date, start time, and finish time in 'record'
            task_date = pd.to_datetime(record.get('day')).strftime('%b %Y')  # Extract 'date' from the record
            start_time = record.get('dates', {}).get('start')  # Assuming 'start' time is available
            finish_time = record.get('dates', {}).get('finish')  # Assuming 'finish' time is available

            # Filter for categories 'On Hold' or 'Current Timed Projects'
            if category_name in ['On Hold', 'Current Timed Projects']:
                formatted_date = pd.to_datetime(trackedfrom).strftime('%b %Y')

                # Create a new entry for the project if it doesn't exist
                if project_name not in project_data:
                    project_data[project_name] = []

                project_data[project_name].append({
                    'Project Number': project_number,
                    'Project Name' : project_name,
                    'Category' : category_name,
                    'Task Name': task_name,
                    'Timer': task_date,
                    'Staff': contact_name,
                    'Start': start_time,
                    'Finish': finish_time,
                    f'Total Time Spent: {formatted_date}': convert_time(time_spent),
                })

    print("Collected project details...")
    return project_data

# Main function to write separate Excel files per project
def main():

    # Process time for all staff contacts
    project_data = process_time_per_contact(trackedfrom, trackedto)

    # Create output directory if it doesn't exist
    output_dir = 'output/projects'
    os.makedirs(output_dir, exist_ok=True)

    # Write each project's data to a separate Excel file
    for project_name, records in project_data.items():
        df = pd.DataFrame(records)
        # Replace invalid characters in project names that can't be used in file names
        safe_project_name = "".join([c if c.isalnum() or c in (' ', '-', '_') else '_' for c in project_name])
        formatted_date = pd.to_datetime(trackedfrom).strftime('%b %Y')
        project_number = records[0]['Project Number']
        excel_file = f'{output_dir}/{project_number} - {safe_project_name} {formatted_date} Timesheet.xlsx'

        # Write DataFrame to Excel
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']  # Access the sheet

            # Adjust the width of each column to fit the content
            for col_idx, col in enumerate(df.columns, 1):
                max_length = max(df[col].astype(str).map(len).max(), len(col))  # Find the max length in the column
                worksheet.column_dimensions[get_column_letter(col_idx)].width = max_length + 2  # Add a little padding

        print(f"Time data for project '{project_name}' saved to {excel_file}")

if __name__ == "__main__":
    main()
