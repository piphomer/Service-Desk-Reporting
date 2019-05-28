#Python 3

from jira import JIRA
import csv
import sys
from datetime import datetime as dt
import os

board_list =['BMT']

#Auth
jira_username = os.environ['JIRA_USERNAME']
jira_api_token = 'LBTfuc89zOImffuH61Hm3FB3'

if __name__ == "__main__":

    run_date = dt.strftime(dt.now(), '%y%m%d')

    print(run_date)

    # Connect to BBOXX Jira server
    jira = JIRA('https://bboxxltd.atlassian.net', basic_auth=(jira_username, jira_api_token))


    for board_id in board_list:

        search_string = 'project = ' + board_id

        # print(search_string)

        # Retrieve all tickets so we can count them.
        # Note: we can only query 100 at a time so we will need to paginate in a later step
        all_tix = jira.search_issues(search_string)

        tix_count = all_tix.total # Count the tickets via the .total attribute

        page_qty = tix_count // 100 + 1 # Calculate how many pages of tickets there are

        print("Total number of tickets: ",tix_count)
        print("Number of pages: ", page_qty)

        output_list = []

        # Loop through the number of pages we need to gather all issues
        for page in range(page_qty):

            print("Page: ", page + 1)

            starting_issue = page * 100

            tix = jira.search_issues(search_string, startAt = starting_issue, maxResults= 100)
            
            for issue in tix:

                issue_key = issue.key
                
                try:
                    issuetype = issue.fields.issuetype.name
                except:
                    issuetype = "-"

                try:
                    reporter = issue.fields.reporter.displayName
                except:
                    reporter = "-"

                try:
                    priority = issue.fields.priority.name
                except:
                    priority = "-"

                try:
                    root_cause = issue.fields.customfield_11446[0].value
                except:
                    root_cause = "-"

                try:
                    resolution_bf = issue.fields.customfield_11447.value
                except:
                    resolution_bf = "-"
                try:
                    status = issue.fields.status.name
                except:
                    status = "-"

                try:
                    created = issue.fields.created
                except:
                    created = "-"

                try:
                    assignee = issue.fields.assignee.displayName
                except:
                    assignee = "-"

                labels_list = []

                for i in range(5):
                    try:
                        labels_list.append(issue.raw["fields"]["labels"][i])
                    except:
                        labels_list.append("-")

                if issue.raw['fields']['customfield_10500'] != "{}":
                    development = "True"
                else: development = "False"

                summary = issue.fields.summary

                try:
                    assignee = issue.fields.assignee.displayName
                except:
                    assignee = "-"

                
                this_list = [issue_key,issuetype,reporter,assignee,priority,root_cause,resolution_bf,status,created,
                                labels_list[0],labels_list[1],labels_list[2],labels_list[3],labels_list[4],development,summary]

                # print(this_list)

                output_list.append(this_list)
            

        header_list = ["Issue","Issue Type","Reporter","Assignee","Priority","Root Cause","Resolution (BF)",
                                        "Status","Created","Label 1","Label 2","Label 3","Label 4","Label 5","Development","Summary"]
        

        #Write all metrics to csv file
        fname = "../outputs/{}_board_dump_{}.csv".format(run_date,board_id)

        with open(fname, 'w', encoding='utf-8', newline='') as csvfile:
            
            print("Writing .csv file...")

            writer = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC)
            writer.writerow(header_list)
            writer.writerows(output_list)

