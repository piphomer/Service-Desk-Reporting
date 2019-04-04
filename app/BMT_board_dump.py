from jira import JIRA
import csv
import sys
from datetime import datetime as dt
import pprint
import pandas as pd

board_list =['BMT']

if __name__ == "__main__":

    run_date = dt.strftime(dt.now(), '%y%m%d')

    print(run_date)

    # Connect to BBOXX Jira server
    jira = JIRA('https://bboxxltd.atlassian.net', basic_auth=('p.homer@bboxx.co.uk', 'S0trwuh#1'))

    #choose which desk to run
    # PS = Product Support
    # CMS = Customer Management Support
    # ES = Expansion Support

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

<<<<<<< HEAD:app/BMT_board_dump.py
                issue_key = issue.key
                
                try:
                    issuetype = issue.fields.issuetype.name
=======
                issue_key = issue.key.encode('utf-8')
                try:
                    issuetype = issue.fields.issuetype.name.encode('utf-8')
>>>>>>> parent of 8f64e57... Added Label 1 - 5 fields for BMT dump:BMT_board_dump.py
                except:
                    issuetype = "-"
                try:
<<<<<<< HEAD:app/BMT_board_dump.py
                    reporter = issue.fields.reporter.displayName
=======
                    reporter = issue.fields.reporter.displayName.encode('utf-8')
>>>>>>> parent of 8f64e57... Added Label 1 - 5 fields for BMT dump:BMT_board_dump.py
                except:
                    reporter = "-"
                try:
<<<<<<< HEAD:app/BMT_board_dump.py
                    priority = issue.fields.priority.name
=======
                    priority = issue.fields.priority.name.encode('utf-8')
>>>>>>> parent of 8f64e57... Added Label 1 - 5 fields for BMT dump:BMT_board_dump.py
                except:
                    priority = "-"
                try:
<<<<<<< HEAD:app/BMT_board_dump.py
                    root_cause = issue.fields.customfield_11446[0].value
=======
                    root_cause = issue.fields.customfield_11446[0].value.encode('utf-8')
>>>>>>> parent of 8f64e57... Added Label 1 - 5 fields for BMT dump:BMT_board_dump.py
                except:
                    root_cause = "-"
                try:
<<<<<<< HEAD:app/BMT_board_dump.py
                    resolution_bf = issue.fields.customfield_11447.value
=======
                    resolution_bf = issue.fields.customfield_11447.value.encode('utf-8')
>>>>>>> parent of 8f64e57... Added Label 1 - 5 fields for BMT dump:BMT_board_dump.py
                except:
                    resolution_bf = "-"
                try:
<<<<<<< HEAD:app/BMT_board_dump.py
                    status = issue.fields.status.name
=======
                    status = issue.fields.status.name.encode('utf-8')
>>>>>>> parent of 8f64e57... Added Label 1 - 5 fields for BMT dump:BMT_board_dump.py
                except:
                    status = "-"
                try:
<<<<<<< HEAD:app/BMT_board_dump.py
                    created = issue.fields.created
=======
                    created = issue.fields.created.encode('utf-8')
>>>>>>> parent of 8f64e57... Added Label 1 - 5 fields for BMT dump:BMT_board_dump.py
                except:
                    created = "-"

                try:
<<<<<<< HEAD:app/BMT_board_dump.py
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
=======
                    assignee = issue.fields.assignee.displayName.encode('utf-8')
                except:
                    assignee = "-"

                this_list = [issue_key,issuetype,reporter,assignee,priority,root_cause,resolution_bf,status,created]
>>>>>>> parent of 8f64e57... Added Label 1 - 5 fields for BMT dump:BMT_board_dump.py

                this_list = [issue_key,issuetype,reporter,assignee,priority,root_cause,resolution_bf,status,created,
                                labels_list[0],labels_list[1],labels_list[2],labels_list[3],labels_list[4],development,summary]

                # print(this_list)

                output_list.append(this_list)
            

<<<<<<< HEAD:app/BMT_board_dump.py
        header_list = ["Issue","Issue Type","Reporter","Assignee","Priority","Root Cause","Resolution (BF)",
                                        "Status","Created","Label 1","Label 2","Label 3","Label 4","Label 5","Development","Summary"]
        
=======
        output_df = pd.DataFrame(output_list)
        output_df.columns = ["Issue","Issue Type","Reporter","Assignee","Priority",
                                "Root Cause","Resolution (BF)","Status","Created"]
        output_df.set_index("Issue", inplace=True)

        print(output_df)
>>>>>>> parent of 8f64e57... Added Label 1 - 5 fields for BMT dump:BMT_board_dump.py

        #Write all metrics to csv file
        fname = "../outputs/{}_board_dump_{}.csv".format(run_date,board_id)

<<<<<<< HEAD:app/BMT_board_dump.py
        with open(fname, 'w', encoding='utf-8', newline='') as csvfile:
            
            print("Writing .csv file...")

            writer = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC)
            writer.writerow(header_list)
            writer.writerows(output_list)

=======
        output_df.to_csv(fname)
>>>>>>> parent of 8f64e57... Added Label 1 - 5 fields for BMT dump:BMT_board_dump.py
