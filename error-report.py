import os
import shutil
import zipfile
import smtplib
from email.mime.text import MIMEText
from datetime import datetime
from time import sleep
import requests
import uuid
import configparser

# Load configuration from the file
config = configparser.ConfigParser()
config.read("error-report-config.ini")
#print(config.sections())
#agfa variables

error_report_repo = config.get("Agfa", "error_report_repo")
#test source folder, for prod change back to "E:\\Agfa\\error-report"
source_folder = config.get("Agfa", "source_folder")
search_term = config.get("Agfa", "search_term")

#email variables
smtp_server = config.get("Email", "smtp_server")
smtp_port = config.get("Email", "smtp_port")
smtp_username = config.get("Email", "smtp_username")
smtp_password = config.get("Email", "smtp_password")
smtp_from_domain = config.get("Email", "smtp_from_domain")
smtp_from = f"{os.environ['COMPUTERNAME']}@{smtp_from_domain}"
smtp_recipients_string = config.get("Email", "smtp_recipients")
smtp_recipients = smtp_recipients_string.split(",")

#service now variables
service_now_instance = config.get("ServiceNow", "instance")
service_now_table = config.get("ServiceNow", "table")
service_now_api_user = config.get("ServiceNow", "api_user")
service_now_api_password = config.get("ServiceNow", "api_password")
ticket_type = config.get("ServiceNow", "ticket_type")
configuration_item = config.get("ServiceNow", "configuration_item")
assignment_group = config.get("ServiceNow", "assignment_group")
assignee = config.get("ServiceNow", "assignee")
business_hours_start_time = config.get("ServiceNow", "business_hours_start_time")
business_hours_end_time = config.get("ServiceNow", "business_hours_end_time")
after_hours_urgency = config.get("ServiceNow", "after_hours_urgency")
after_hours_impact = config.get("ServiceNow", "after_hours_impact")
business_hours_urgency = config.get("ServiceNow", "business_hours_urgency")
business_hours_impact = config.get("ServiceNow", "business_hours_impact")

#excluded items variables
excluded_computer_names_path = config.get("Excludeditems", "excluded_computer_names_path")
excluded_user_codes_path = config.get("Excludeditems", "excluded_user_codes_path")



# Get the current time and day of the week
current_time = datetime.now().time()
current_day = datetime.now().weekday()

# Define business hours
business_hours_start = datetime.strptime(business_hours_start_time, "%H:%M:%S").time()
business_hours_end = datetime.strptime(business_hours_end_time, "%H:%M:%S").time()

# Set default values
urgency = after_hours_urgency  # Default value for after hours and weekends
impact = after_hours_urgency   # Default value for after hours and weekends       

# Check if it's business hours
if business_hours_start <= current_time <= business_hours_end and current_day < 5:  # Monday to Friday
    urgency = business_hours_urgency
    impact = business_hours_impact


def read_excluded_values(file_path):
    with open(file_path, "r") as file:
        return [line.strip() for line in file.readlines()]
        
# Load excluded values from text files
excluded_computer_names = read_excluded_values(excluded_computer_names_path)
excluded_user_codes = read_excluded_values(excluded_user_codes_path)        


def send_email(smtp_recipients, subject, body):
    msg = MIMEText(body)
    msg["From"] = smtp_from
    msg["To"] = ", ".join(smtp_recipients)  # Join smtp_recipients with a comma and space
    msg["Subject"] = subject

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.sendmail(smtp_from, smtp_recipients, msg.as_string())
        server.quit()
        print(f"Email sent to {', '.join(smtp_recipients)}")
    except Exception as e:
        print(f"Email sending failed to {', '.join(smtp_recipients)}: {e}")

def create_service_now_incident(summary, description, affected_user_id, configuration_item, external_unique_id, urgency, impact, device_name, ticket_type):
    incident_api_url = f"https://{service_now_instance}/api/now/table/{service_now_table}"

    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

    payload = {
        "u_short_description": summary,
        "u_description": description,
        "u_affected_user_id": affected_user_id,
        "u_configuration_item": configuration_item,
        "u_external_unique_id": external_unique_id,
        "u_urgency": urgency,
        "u_impact": impact,
        "u_type": ticket_type,
        "u_assignment_group": assignment_group,
    }

    try:
        print("Incident Creation Payload:", payload)  # Print payload for debugging
        response = requests.post(
            incident_api_url,
            headers=headers,
            auth=(service_now_api_user, service_now_api_password),      
            json=payload,
        )

        print("Incident Creation Response Status Code:", response.status_code)  # Print status code for debugging
        print("Incident Creation Response Content:", response.text)  # Print response content for debugging

        if response.status_code == 201:
            incident_number = response.json().get("result", {}).get("u_task_string")
            sys_id = response.json().get('result', {}).get('u_task', {}).get('value')
            print(f"ServiceNow incident created successfully: {incident_number} {sys_id}")
            return incident_number, sys_id
        else:
            print(f"Failed to create ServiceNow incident. Response: {response.text}")

    except requests.exceptions.RequestException as e:
        print(f"An error occurred while creating ServiceNow incident: {e}")

    return None, None



def attach_file_to_incident(incident_number, file_path):
    

    attachment_api_url = f"https://{service_now_instance}/api/now/attachment/upload"

    headers = {
        "Accept": "application/json",
    }

    data = {
        "table_name": service_now_table,
        "table_sys_id": incident_number,
    }

    files = {
        'file': (os.path.basename(file_path), open(file_path, 'rb')),
    }

    try:
        print("Sending attachment request...")
        print("Data:", data)
        print("Files:", files)
        attachment_response = requests.post(
            attachment_api_url,
            headers=headers,
            auth=(service_now_api_user, service_now_api_password),
            data=data,
            files=files,
        )

        print("Attachment response status code:", attachment_response.status_code)
        print("Attachment response text:", attachment_response.text)

        if attachment_response.status_code == 201:
            print("File attached to ServiceNow incident successfully")
        else:
            print(f"Failed to attach file to ServiceNow incident. Response: {attachment_response.text}")

    except requests.exceptions.RequestException as e:
        print(f"An error occurred while attaching the file to ServiceNow incident: {e}")



if not os.path.exists(error_report_repo):
    os.makedirs(error_report_repo)

for root, _, files in os.walk(source_folder):
    for file in files:
        if search_term in file:
            source_path = os.path.join(root, file)
            relative_path = os.path.relpath(source_path, source_folder)
            destination_path = os.path.join(error_report_repo, relative_path)
            
            destination_directory = os.path.dirname(destination_path)
            os.makedirs(destination_directory, exist_ok=True)
            if os.path.exists(destination_path):
                continue

            original_timestamp = os.path.getmtime(source_path)  # Get the original file's timestamp
            shutil.copy2(source_path, destination_path)  # Use shutil.copy2() to preserve timestamps
            print(f"working on {source_path}")

            with zipfile.ZipFile(destination_path, "r") as zip_ref:
                user_code = None
                comment_content = None
                for entry in zip_ref.namelist():
                    if "comment.txt" in entry:
                        with zip_ref.open(entry) as comment_file:
                            comment_content = comment_file.read().decode("utf-8")
                            break

                    if entry.startswith("logs/") and "agility" in entry:
                        with zip_ref.open(entry) as log_file:
                            for line in log_file:
                                line = line.decode("utf-8")
                                if "userCode=" in line:
                                    user_code = line.split("userCode=")[1].split("@")[0]
                                    break

                computer_name = os.path.dirname(relative_path).lstrip("\\")
                local_time = datetime.fromtimestamp(original_timestamp)
                local_time_str = local_time.strftime('%Y-%m-%d %H:%M:%S')
                subject = f"Client Error Report for {computer_name} at {local_time_str} (Ticket Creation Failure)"
                body = f"Content of comment.txt:\n{comment_content}\nUserID: {user_code}\nWorkstation: {computer_name}"
                incident_summary = f"Client Error Report for {computer_name} at {local_time_str}"
                incident_description = body
                affected_user_id = user_code
                device_name = computer_name
                external_unique_id = str(uuid.uuid4())
                
                # Check if the item should be excluded
                if affected_user_id and affected_user_id.lower() in [code.lower() for code in excluded_user_codes] or \
                   computer_name.lower() in [name.lower() for name in excluded_computer_names]:
                    print(f"Skipping ServiceNow processing for excluded computer_name or user_code: {computer_name} - {affected_user_id}")
                    subject = f"Client Error Report for {computer_name} at {local_time_str} (Ticket Exclusion)"
                    send_email(smtp_recipients, subject, body)
                    continue                  

                # Create ServiceNow incident and get the incident number
                incident_number, sys_id = create_service_now_incident(
                    incident_summary, incident_description,
                    affected_user_id, configuration_item, external_unique_id,
                    urgency, impact, device_name, ticket_type
                )

                # Attach the zip file to the ServiceNow incident
                if incident_number and sys_id:
                    zip_file_path = destination_path
                    attach_file_to_incident(sys_id, zip_file_path)
                    subject = f"Client Error Report for {computer_name} at {local_time_str} Ticket: {incident_number}"
                
                
                 #send email
                send_email(smtp_recipients, subject, body)
                
                sleep(1)  # Introduce a delay of 1 second before working on next error report
