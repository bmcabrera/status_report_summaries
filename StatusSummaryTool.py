### Status Summary Tool ###
import win32com.client
import datetime
import re
import pandas as pd
import sys

def check_date_args(start_date_arg, end_date_arg):
    try:
        START_DATE = datetime.datetime.strptime(start_date_arg, '%m/%d/%Y')
        END_DATE = datetime.datetime.strptime(end_date_arg, '%m/%d/%Y')
    except ValueError:
        print('Invalid date format. Please enter a date in the format MM/DD/YYYY.')
        sys.exit(1)
    return START_DATE, END_DATE
    
def date_range_filter(emails, START_DATE, END_DATE):
    """Filters a collection of Outlook emails by a date range.
    
    Args:
        emails: A collection of Outlook emails.
        START_DATE: The start of the date range to filter by.
        END_DATE: The end of the date range to filter by.
        
    Returns:
        A collection of Outlook emails that fall within the specified date range.
    """
    print("Filtering down to only emails from " + START_DATE.strftime('%m/%d/%Y') + " to " + END_DATE.strftime('%m/%d/%Y') + "...")
    start_date_str = START_DATE.strftime('%m/%d/%Y %H:%M %p')
    end_date_str = END_DATE.strftime('%m/%d/%Y %H:%M %p')

    date_restriction = "[ReceivedTime] >= '" + start_date_str + "' AND [ReceivedTime] <= '" + end_date_str + "'"
    filtered_emails = emails.Restrict(date_restriction)
    return filtered_emails

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

def find_member_reports(text, team_regex_pattern):
    """Searches a string for all non-overlapping occurrences of a regular expression pattern and returns an iterator of match objects.

    Args: 
        text: The string to search.
        team_regex_pattern: The regular expression pattern to search for.

    Returns:
        A list of strings for each member's report in the email.
    """
    matches = re.finditer(team_regex_pattern, text)
    return matches

def main(start_date_arg, end_date_arg):        
    ##################### User Variables #####################

    # user's email 
    # Example: 'fname.lname@lexisnexisrisk.com'
    USER_EMAIL = 'bryan.cabrera@lexisnexisrisk.com'

    #get_report_for = []

    # email folder to search
    # Example: 'Inbox'
    EMAIL_FOLDER = 'Inbox'

    # subject line to search for
    # Example: 'Group Status'
    EMAIL_SUBJECT = 'Group Status'

    # date range to use to parse emails
    # Example:
    START_DATE = check_date_args(start_date_arg, end_date_arg)[0]
    END_DATE = check_date_args(start_date_arg, end_date_arg)[1]
    '''
    START_DATE = datetime.datetime(2025, 2, 10)
    END_DATE = datetime.datetime(2025, 2, 18)
    '''
    ##########################################################

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    USER_EMAIL = outlook.Folders(USER_EMAIL)
    EMAIL_FOLDER = USER_EMAIL.Folders(EMAIL_FOLDER)

    # get emails in <EMAIL_FOLDER> and sort by received time
    emails = EMAIL_FOLDER.Items
    emails.Sort('[ReceivedTime]', True)

    # historical list of team members used for parsing the email body
    team = ['Attila', 'Bob', 'Bryan', 'Chris', 'Godji', 'Greg', 'Jim', 'Ken', 'Kunal', 'Michael', 'Ming', 'Richard']
    team_regex_pattern = ':|'.join(team)+':'
    final_report_df = pd.DataFrame(columns=['Sent_Date', 'Subject_Date', 'Name', 'Report_Summary'])



    # parse all emails or emails within a specific date range
    if not (START_DATE or END_DATE):
        print(f'Searching...\n \tEmail: {USER_EMAIL}' + f'\n\tFolder: {EMAIL_FOLDER}' + f'\n\tSubject: {EMAIL_SUBJECT}' + f'\n\tDate Range: ALL IN "{EMAIL_FOLDER}" FOLDER\n')
        emails_to_parse = emails
    else:
        print('Searching...\n \tEmail: ' + str(USER_EMAIL) + '\n\tFolder: ' + str(EMAIL_FOLDER) + '\n\tSubject: ' + str(EMAIL_SUBJECT) + '\n\tDate Range: ' + START_DATE.strftime('%m/%d/%Y') + ' to ' + END_DATE.strftime('%m/%d/%Y') + '\n')
        emails_to_parse = date_range_filter(emails, START_DATE, END_DATE)

    # iterate through narrowed down emails, search if <EMAIL_SUBJECT> (Group Status) is found in each email's subject line, and parse the body of those emails
    for email in emails_to_parse:
        if EMAIL_SUBJECT in email.subject:
            ### PARSING EMAIL BODY ###

            # search each email body, find team members in our historical_team_members list/pattern, and add them to the team_members_found list
            matches = find_member_reports(email.body, team_regex_pattern)
            team_members_found = [match.group() for match in matches]
            #print(email.subject)
            #print(team_members_found)

            with open("my_text_file.txt", "w", encoding="utf-8") as file:
                # Write the string to the file.
                file.write(str(email.htmlbody))
        
            # split the body of each email using each team member's name (in the team_members_found list) as a delimiter creating a list (member_reports) of each report
            member_reports = split_string(email.body.replace('\r', '').replace('\n', '').replace('\t', ' '), team_members_found)
            member_reports = list(filter(None, member_reports))
            #print(member_reports)
            
            # extract the date from the subject line
            extracted_subject_date = datetime.datetime.strptime(re.findall(r'\d+/\d+/\d+', email.subject)[0], '%m/%d/%Y').date()
            
            
            # combine team members to their parsed report into a dataframe
            for member in team_members_found:
                final_report_df.loc[len(final_report_df)] = [email.ReceivedTime.date(), extracted_subject_date, member, member_reports[team_members_found.index(member)]]
            

    ### CLEAN FINAL DATAFRAME ###
    # remove colons from name column
    final_report_df['Name'] = final_report_df['Name'].str.replace(':', '')
    final_report_df = final_report_df.sort_values(by=['Sent_Date', 'Name'], ascending=[False, True])
    final_report_df.to_csv('StatusReportSummary_'+str(datetime.datetime.now().strftime('%Y-%m-%d_%H-%M'))+'.csv', index=False)



if __name__=="__main__":
    main(sys.argv[1], sys.argv[2])