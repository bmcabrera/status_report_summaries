### Status Summary Tool ###
import win32com.client
import datetime
import re
import os

# Example: 'fname.lname@lexisnexisrisk.com'
user_email = ''
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# User's email
user_email = outlook.Folders(user_email)
print('Searching Email: ' + str(user_email) + '...')

# Email's folder <inbox> to search through
inbox = user_email.Folders("Inbox")

# Get emails in inbox and sort by received time
emails = inbox.Items
emails.Sort("[ReceivedTime]", True)

# List of team members
# Example: ['Name1', 'Name2', 'Name3', ...]
team = []
delim_team = [s + ':' for s in team]
seperated_status_list = []

def split_string(text, delimiters):
    """Splits a string by multiple delimiters.

    Args:
        text: The string to split.
        delimiters: A list of delimiter strings.

    Returns:
        A list of strings resulting from the split.
    """
    regex_pattern = "|".join(map(re.escape, delimiters))
    return re.split(regex_pattern, text)

# Iterating through all the emails in the inbox and searching for those with a subject that starts with "Group Status" to parse
for email in emails:
    # startswith("Group Status") to find all group status emails in inbox
    if email.subject.startswith("Group Status"):
        term_size = os.get_terminal_size()
        print('\n')
        print(('***'+email.subject+'***').center(term_size.columns))
        for team_member_name in team:
            print('\t'+team_member_name)
        print('\n')
        print('=' * term_size.columns)

        # Splitting the body of the email by new lines
        results = split_string(email.body, delim_team)
        results = list(filter(None, results))
        for i in range(len(results)):
            seperated_status_list.append(delim_team[i] + '\n' + results[i])
            print(seperated_status_list[i])
            print('=' * term_size.columns)

