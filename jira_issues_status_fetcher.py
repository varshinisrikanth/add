"""
This script interacts with the Jira REST API to fetch and process issue data based on a JQL query.

The primary functionality includes:
- Fetching issues from Jira using a specified JQL query and batch size.
- Extracting detailed information about the issues status transitions, fix versions,sprint names.
- Processing and formatting the extracted data into a structured format for analysis.
- Exporting the processed issue data to both JSON and Excel formats.

Modules used:
- `requests`: For making HTTP requests to the Jira API.
- `json`: For parsing and generating JSON data.
- `logging`: For logging the script's activity and errors.
- `argparse`: For handling command-line arguments.
- `re`: For extracting specific patterns (e.g., sprint names) from Jira's response data.
- `pandas`: For handling and exporting data to Excel format.
- `urllib3`: For managing SSL warnings in HTTP requests.

"""
#jira_issues_status_fetcher
import json
import logging
import argparse
import re
import os
import requests
import urllib3
import pandas as pd

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Setup logging to capture the process details
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class JiraClient:
    """
    A client class to interact with Jira's REST API to fetch issue data and transitions.

    This class provides methods to:
    1. Fetch issues from Jira using JQL (Jira Query Language).
    2. Fetch available transitions for a specific Jira issue.
    
    Attributes:
        jira_url (str): The URL of the Jira instance.
        auth (tuple): A tuple containing the username and password for Jira authentication.
        headers (dict): The HTTP headers used for API requests.

    Methods:
        fetch_issues(jql_query, batch_size):
            Fetches issues from Jira based on the provided JQL query and batch size.
        fetch_issue_transitions(issue_key):
            Fetches the available transitions for a specific Jira issue.
    """
    def __init__(self, jira_url, username, password):
        self.jira_url = jira_url
        self.auth = (username, password)
        self.headers = {'Content-Type': 'application/json'}

    def fetch_issues(self, jql_query, batch_size):
        """Fetch issues from Jira based on JQL query."""
        payload = json.dumps({
            'jql': jql_query,
            'maxResults': batch_size,
            'fields': ['key', 'issuetype', 'status', 'timetracking', 'subtasks',
                       'tasks', 'fixVersions', 'customfield_10583'],
            'expand': ['changelog']
        })
        response = requests.post(
            f'{self.jira_url}/rest/api/2/search',
            headers=self.headers,
            data=payload,
            auth=self.auth,
            verify=False,
            timeout=30
        )
        if response.status_code == 200:
            return response.json()
        logging.error("Error fetching issues: %s", response.status_code)
        return None

    def fetch_issue_transitions(self, issue_key):
        """Fetch available transitions for a specific Jira issue."""
        transitions_url = f"{self.jira_url}/rest/api/2/issue/{issue_key}/transitions"
        response = requests.get(
            transitions_url,
            headers=self.headers,
            auth=self.auth,
            verify=False,
            timeout=30
        )

        if response.status_code == 200:
            return response.json().get('transitions', [])
        logging.error("Error fetching transitions for issue %s: %s",issue_key,response.status_code)
        return []


def extract_changelog_transitions(changelog):
    """Extract transitions from the changelog and format them as 'from state -> to state'."""
    transitions = []
    for history in changelog:
        for item in history.get('items', []):
            if item['field'] == 'status':
                from_state = item.get('fromString', 'Unknown')
                to_state = item.get('toString', 'Unknown')
                transitions.append(f"{from_state} -> {to_state}")
    return ', '.join(transitions) if transitions else "No transitions"


def extract_fix_versions(issue):
    """Extract and format the fix versions for a given issue."""
    fix_versions = issue['fields'].get('fixVersions', [])
    fix_versions_list = [
        version.get('name', 'None') for version in fix_versions
    ] if fix_versions else ['None']
    return ', '.join(fix_versions_list)


def extract_sprint_names(issue):
    """Extract the sprint names from the issue and return as a comma-separated string."""
    sprint_field = issue['fields'].get('customfield_10583', None)  # Sprint custom field
    sprint_names = []

    # Log the type and contents of sprint_field for debugging purposes
    if sprint_field:
        for sprint_info in sprint_field:
            # Use regular expression to extract the 'name' field from the string representation
            match = re.search(r'name=([A-Za-z0-9\s\-]+)', sprint_info)
            if match:
                sprint_names.append(match.group(1))  # Extracted sprint name

    # If no sprint names are found, return 'No Sprint'
    return ', '.join(sprint_names) if sprint_names else 'No Sprint'



def count_open_closed_issues(status, open_count, closed_count):
    """Update the open and closed counts based on issue status."""
    if 'open' in status:
        open_count += 1
    elif 'closed' in status:
        closed_count += 1
    return open_count, closed_count


def prepare_issue_data(issue, issue_key, changelog_transitions, fix_versions_str, sprint_names_str):
    """Prepare a dictionary of issue data for JSON export."""
    return {
        'Issue Key': issue_key,
        'Issue Type': issue['fields'].get('issuetype', {}).get('name', 'Unknown'),
        'Current Status': issue['fields'].get('status', {}).get('name', 'Unknown').strip().lower(),
        'Estimated Effort': (
            issue['fields']
            .get('timetracking', {})
            .get('originalEstimateSeconds', 0) / 3600),
        'Actual Effort': issue['fields'].get('timetracking', {}).get('timeSpentSeconds', 0) / 3600,
        'Rounded Actual Effort': round(
            issue['fields']
            .get('timetracking', {})
            .get('timeSpentSeconds', 0) / 3600),
        'Transitions': changelog_transitions,
        'Fix Versions': fix_versions_str,
        'Sprint': sprint_names_str  # Add the sprint names to the data
    }


def fetch_and_process_issues(jql_query, batch_size, jira_client):
    """Fetch issues using the provided JQL query and process them."""
    issues_data = jira_client.fetch_issues(jql_query, batch_size)
    if not issues_data:
        logging.error("No data fetched from Jira.")
        return [], 0, 0  # Return empty data and 0 counts
    open_count, closed_count = 0, 0
    issue_data = []
    # Sort issues by issue key numerically
    sorted_issues = sorted(
        issues_data['issues'],
        key=lambda issue: int(issue['key'].split('-')[1])
    )
    for issue in sorted_issues:
        issue_key = issue['key']
        logging.info("Processing issue: %s", issue_key)
        # Extract changelog transitions and fix versions
        changelog_transitions = extract_changelog_transitions(
            issue.get('changelog', {}).get('histories', [])
        )
        fix_versions_str = extract_fix_versions(issue)
        # Fetch sprint name (customfield_10583 is assumed to be the sprint field)
        sprint_names_str = extract_sprint_names(issue)
        # Count open and closed issues based on status
        status = issue['fields'].get('status', {}).get('name', 'Unknown').strip().lower()
        open_count, closed_count = count_open_closed_issues(status, open_count, closed_count)
        # Prepare issue data including sprint name
        issue_data.append(
            prepare_issue_data(
                issue,
                issue_key,
                changelog_transitions,
                fix_versions_str,
                sprint_names_str  # Include sprint name in the data
            )
        )
    return issue_data, open_count, closed_count


def export_to_excel(data, file_path):
    """Export the fetched issues data and summary to an Excel file."""
    # Convert the 'issues' list to a pandas DataFrame
    issues_df = pd.DataFrame(data['issues'])
    # Create a DataFrame for the summary data (open/closed counts)
    summary_df = pd.DataFrame([data['summary']])

    # Write both DataFrames to separate sheets in the same Excel file
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        issues_df.to_excel(writer, sheet_name='Issues', index=False)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)


def main():
    """Main function to orchestrate the process."""
    try:
        # Parse command-line arguments
        parser = argparse.ArgumentParser(description="Fetch Jira issues based on JQL query")
        parser.add_argument('--jira_url', required=True, help="Jira URL")
        parser.add_argument('--jira_username', required=True, help="Jira Username")
        parser.add_argument('--jira_password', required=True, help="Jira Password")
        parser.add_argument('--jql_query', required=True, help="JQL query to fetch issues")
        parser.add_argument('--batch_size', type=int, default=100, help="Number of issues")
        args = parser.parse_args()

        # Instantiate JiraClient
        jira_client = JiraClient(
        args.jira_url.strip(),
        args.jira_username.strip(),
        args.jira_password.strip()
        )


        # Fetch and process issues
        logging.info("Fetching issues from Jira using the JQL query")
        issue_data, open_count, closed_count = fetch_and_process_issues(
        args.jql_query.strip(),
        args.batch_size,
        jira_client
        )

        # If no issues were fetched, return early
        if not issue_data:
            logging.error("No issues found.")
            return

        # Prepare summary data
        summary_data = {
        'Open Issues': open_count,
        'Closed Issues': closed_count
        }

        # Define file paths in the current script directory
        script_directory = os.path.dirname(os.path.abspath(__file__))
        json_file_path = os.path.join(script_directory, 'issues_data_with_transitions.json')
        excel_file_path = os.path.join(script_directory, 'issues_data_with_transitions.xlsx')

        # Save the fetched issue data to a JSON file
        with open(json_file_path, 'w', encoding='utf-8') as json_file:
            json.dump({
                'issues': issue_data,
                'summary': summary_data
            }, json_file, ensure_ascii=False, indent=4)
        logging.info("Fetched issue data has been saved to %s", json_file_path)

        # Now export the data to Excel
        export_to_excel({
            'issues': issue_data,
            'summary': summary_data
        }, excel_file_path)
        logging.info("Data successfully exported to %s", excel_file_path)

    except ValueError as error:
        logging.error("An error occurred: %s", str(error))


if __name__ == '__main__':
    main()
