from jira import JIRA
import csv
import sys
import os
import requests
from datetime import datetime as dt
import unicodedata
from unidecode import unidecode



# from office365.runtime.auth.authentication_context import AuthenticationContext
# from office365.sharepoint.client_context import ClientContext
# from office365.sharepoint.file_creation_information import FileCreationInformation
# from office365.runtime.utilities.request_options import RequestOptions

#List of service desks to iterate through
<<<<<<< HEAD:app/jira_service_desk_dump.py
service_desk_list = ['PS','CMS']
# service_desk_list = ['CMS']
=======
service_desk_list = ['CMS','PS', 'ES']
service_desk_list = ['CMS']
>>>>>>> parent of 8f64e57... Added Label 1 - 5 fields for BMT dump:jira_service_desk_dump.py

#Auth
jira_username = os.environ['JIRA_USERNAME']
jira_password = os.environ['JIRA_PASSWORD']

sharepoint_username = os.environ['SHAREPOINT_USERNAME']
sharepoint_password = os.environ['SHAREPOINT_PASSWORD']
# sharepoint_url = os.environ['SHAREPOINT_URL']

sharepoint_url = 'https://bboxxeng.sharepoint.com'



####################################################################################################
# Set up Sharepoint functions
####################################################################################################

class Sharepoint():

    """ Sharepoint Connection """
    def __init__(self, sharepoint_username, sharepoint_password, sharepoint_url):
        self.sharepoint_username = sharepoint_username
        self.sharepoint_password = sharepoint_password
        self.sharepoint_url = sharepoint_url

    def authorise_sharepoint(self, url=None):
        if not url:
            raise Exception("Error","No Sharepoint URL is defined in config.")
        ctx_auth = AuthenticationContext(url=url)
        self.sp_url = url
        if ctx_auth.acquire_token_for_user(username=self.sharepoint_username,
                                           password=self.sharepoint_password):
            ctx = ClientContext(url, ctx_auth)
            return ctx
        else:
            print("error in auth")
        return False

    def get_folder_by_server_relative_url(self,path):
        if path[0] == '/':
            path = path[1:]
        if path[-1] == '/':
            path = path[:-1]
        folder_paths = path.split('/')
        url = self.sp_url
        if not url:
            raise Exception("Error","No Sharepoint URL is defined in config.")
        if url[-1] != '/':
            url += '/'
        url = url + folder_paths[0] + '/' + folder_paths[1]
        self.ctx = self.authorise_sharepoint(url)
        folder_names = path.split('/')
        list_obj = self.ctx.web.lists.get_by_title(folder_paths[2])
        folder = list_obj.root_folder
        if len(folder_names[2:]) > 1:
             for folder_name in folder_names[3:]:
                # Get destination sharepoint folder
                folder = folder.folders.get_by_url(folder_name)
        return folder

    def upload_file(self,path,file_content,file_name,overwrite = True):
        # Find Folder

        folder_object = self.get_folder_by_server_relative_url(path)

        self.ctx.load(folder_object)
        self.ctx.execute_query()

        # Create File Object
        info = FileCreationInformation()
        info.url = file_name
        info.content = file_content
        info.overwrite = overwrite

        full_url = "{0}/Files/add(url='{1}', overwrite=true)".format(folder_object.url, file_name)


        options = RequestOptions(full_url)
        self.ctx.authenticate_request(options)
        self.ctx.ensure_form_digest(options)

        # Upload File

        # The reason we do this directly is because the
        # `request.execute_query_direct` call wants to send JSON, but that
        # doesn't work if you want to upload, eg, an XLSX file (Requests
        # just falls over trying to decode the contents as text).
        print(('Uploading File "%s" to sharepoint folder "%s" ...' % (info.url, folder_object)))
        print(full_url)
        response = requests.post(
            url=full_url, data=file_content, headers=options.headers, auth=options.auth,
        )
        if response.status_code not in [200,201]:
            raise Exception(response.text, response.status_code)
        print('Done.')



#     def delete_file(self, path):
#         print("Deleting file")
#         split_path = path.rsplit('/', 1)
#         folder_path, filename = split_path[0], split_path[1]
#         folder = self.get_folder_by_server_relative_url(folder_path)
#         self.ctx.load(folder)
#         self.ctx.execute_query()

#         files = folder.files
#         self.ctx.load(files)
#         self.ctx.execute_query()

#         print(files)

#         for cur_file in files:
#             print(cur_file.properties["Name"])
#             if cur_file.properties["Name"] == filename:
#                 full_url = "{0}/$value".format(cur_file.url)
#                 options = RequestOptions(full_url)
#                 self.ctx.authenticate_request(options)
#                 self.ctx.ensure_form_digest(options)
#                 response = requests.delete(url=cur_file.url, headers=options.headers, auth=options.auth)

#                 print(response.status_code)

#                 if response.status_code == 404:
#                     raise SharepointFileNotFound()

#                 elif response.status_code != 200:
#                     raise SharepointDeleteFailed()

#         return


def read_file_as_binary(path):
    with open(path, 'rb') as content_file:
        file_content = content_file.read()
    return file_content


####################################################################################################


if __name__ == "__main__":

    run_date = dt.strftime(dt.now(), '%y%m%d')

    print(run_date)

    # Connect to BBOXX Jira server
    jira = JIRA('https://bboxxltd.atlassian.net', basic_auth=(jira_username, jira_password))

    #Create connection to Sharepoint

    # sp = Sharepoint(sharepoint_username,sharepoint_password, sharepoint_url)

    # sp.authorise_sharepoint(sharepoint_url)
    
    #Move any files in the Active folder to the Archive folder
  
    #Coming soon...


    #Get the tickets

    for desk_id in service_desk_list:

        #Define the header names of each column
        header_list = [
            "Ticket Key",
            "Ticket ID",
            "Issue Type",
            "TTFR Complete?",
            "TTFR Breached?",
            "TTFR (mins)",
            "TTFR Remaining (mins)",
            "TTR Complete?",
            "TTR Breached?",
            "TTR (mins)",
            "TTR Remaining (mins)",
            "Priority",
            "Status",
            "Summary",
            "Created",
            "Reporter",
            "Assignee",
            "Request Type",
            "Affected Product",
            "Organization",
            "Resolution",
            "Resolved",
            "Updated"
        ]

        #CMS desk gets extra columns so add these in
        if desk_id == "CMS":
            header_list.extend([
                "Root Cause",
                "Reason For Breach",
                "Reason for Breach Comment",
                "Resolution ()",
                "Job Type",
                "Department",
                "Job Title",
                "Application 1",
                "Application 2",
                "Application 3",
                "Application 4",
                "Application 5",
                "Application 6",
                "Application 7",
                "Application 8",
                "Application 9",
                "Application 10",
                "CSCC Action",
                "Report Page",
                "Linked Issue",
                "TTFR Orange Completed?",
                "TTFR Orange Breached?",
                "TTFR Orange Elapsed Time (ms)",
                "TTFR Orange Remaining Time (ms)",
                "TTR Orange Completed?",
                "TTR Orange Breached?",
                "TTR Orange Elapsed Time (ms)",
                "TTR Orange Remaining Time (ms)",
                "TTFR Premium Completed?",
                "TTFR Premium Breached?",
                "TTFR Premium Elapsed Time (ms)",
                "TTFR Premium Remaining Time (ms)",
                "TTR Premium Completed?",
                "TTR Premium Breached?",
                "TTR Premium Elapsed Time (ms)",
                "TTR Premium Remaining Time (ms)",
                "TTFR Premium Plus Completed?",
                "TTFR Premium Plus Breached?",
                "TTFR Premium Plus Elapsed Time (ms)",
                "TTFR Premium Plus Remaining Time (ms)",
                "TTR Premium Plus Completed?",
                "TTR Premium Plus Breached?",
                "TTR Premium Plus Elapsed Time (ms)",
                "TTR Premium Plus Remaining Time (ms)",
                "Weekly Status",
                "Last Public Comment Date",
                "ERP Module",
                "Sales Agent Issue",
                "CRM Application",
                "PAYG Issue",
                "Payment Issue",

            ])

        search_string = 'project = ' + desk_id

        print(search_string)

        # Retrieve all tickets so we can count them.
        # Note: we can only query 50 at a time so we will need to paginate in a later step
        all_tix = jira.search_issues(search_string)

        tix_count = all_tix.total # Count the tickets via the .total attribute

        page_qty = tix_count // 100 + 1 # Calculate how many pages of tickets there are

        #page_qty = 1  # Just run first page for test/debug

        print("Total number of tickets: ",tix_count)
        print("Number of pages: ", page_qty)

        output_list = []
        output_list_debug = []

        # Loop through the number of pages we need to gather all issues
        for page in range(page_qty):

            issue_list = []
            issue_list_debug = []

            print("\r" + "Page ", page + 1, " of ", page_qty, end=' ')

            starting_issue = page * 100

            tix = jira.search_issues(search_string, startAt = starting_issue, maxResults= 100)

            for issue in tix:

                # print(issue.key)

                #Issue priority
                try:
                    issue_priority = issue.fields.priority.name
                except:
                    issue_priority = None

                #TTFR
                if issue.raw['fields']['customfield_10806'].get('ongoingCycle'): #if first response has occurred...
                    ttfr_complete = False
                    ttfr_breach = issue.raw['fields']['customfield_10806']['ongoingCycle']['breached']
                    ttfr_time = issue.raw['fields']['customfield_10806']['ongoingCycle']['elapsedTime']['millis'] / 1000 // 60
                    ttfr_remaining = issue.raw['fields']['customfield_10806']['ongoingCycle']['remainingTime']['millis'] / 1000 // 60
                elif issue.raw['fields']['customfield_10806'].get('completedCycles'): #if first response has not occurred
                    ttfr_complete = True
                    ttfr_breach = issue.raw['fields']['customfield_10806']['completedCycles'][0]['breached']
                    ttfr_time = issue.raw['fields']['customfield_10806']['completedCycles'][0]['elapsedTime']['millis'] / 1000 // 60
                    ttfr_remaining = issue.raw['fields']['customfield_10806']['completedCycles'][0]['remainingTime']['millis'] / 1000 // 60
                else:
                    ttfr_complete = ""
                    ttfr_breach = ""
                    ttfr_time = ""
                    ttfr_remaining = ""
                
                #TTR
                if desk_id == "CMS" or desk_id == 'ES':
                    if issue.raw['fields']['customfield_10805'].get('ongoingCycle'): #if first response has occurred...
                        ttr_complete = False
                        ttr_breach = issue.raw['fields']['customfield_10805']['ongoingCycle']['breached']
                        ttr_time = issue.raw['fields']['customfield_10805']['ongoingCycle']['elapsedTime']['millis'] / 1000 // 60
                        ttr_remaining = issue.raw['fields']['customfield_10805']['ongoingCycle']['remainingTime']['millis'] / 1000 // 60
                    elif issue.raw['fields']['customfield_10805'].get('completedCycles'): #if first response has not occurred
                        ttr_complete = True
                        ttr_breach = issue.raw['fields']['customfield_10805']['completedCycles'][-1]['breached']
                        ttr_time = issue.raw['fields']['customfield_10805']['completedCycles'][-1]['elapsedTime']['millis'] / 1000 // 60
                        ttr_remaining = issue.raw['fields']['customfield_10805']['completedCycles'][-1]['remainingTime']['millis'] / 1000 // 60
                    else:
                        ttr_complete = ""
                        ttr_breach = ""
                        ttr_time = ""
                        ttr_remaining = ""
                else:
                    #There is no SLA for TTR on PS desk so the custom field does not get populated with anything useful
                    #So if resolution exists, get the number of days. Otherwise just fill in blanks
                    ttr_complete = ""
                    ttr_breach = ""
                    ttr_remaining = ""
                    if issue.fields.resolutiondate:
                        ttr_time = dt.strptime(issue.fields.resolutiondate[:16], '%Y-%m-%dT%H:%M') - dt.strptime(issue.fields.created[:16], '%Y-%m-%dT%H:%M')
                        ttr_time = ttr_time.days
                    else:
                        ttr_time = ""
                    

                
                #organization
                try:
                    org = issue.raw['fields']['customfield_10700'][0]["name"]
                except:
                    org = "Unknown"

                #Request type
                try:
                    request_type = issue.raw['fields']['customfield_10800']['requestType']['name']
                except:
                    request_type = 'Not specified'

                #Product type
                try:
                    product_type = issue.raw['fields']['customfield_11407'][0]['value']
                except:
                    product_type = 'Not specified'

                #Make the dates Excel-readable
                created = str(issue.fields.created)[:10] + " " + str(issue.fields.created)[11:19]
                resolved = str(issue.fields.resolutiondate)[:10] + " " + str(issue.fields.resolutiondate)[11:19]
                updated = str(issue.fields.updated)[:10] + " " + str(issue.fields.updated)[11:19]
                if ttfr_time != "Ongoing" and ttfr_time != "??":
                    responded = str(ttfr_time)[:10] + " " + str(ttfr_time)[11:19]
                else:
                    responded = "n/a"
                            
                #Assignee
                try:
                    assignee = issue.fields.assignee.displayName
                except:
                    assignee = "None"

                #Resolution
                try:
                    resolution = issue.fields.resolution.name
                except:
                    resolution = "None"

                #Priority
                try:
                    priority = issue.fields.priority.name
                except:
                    priority = "Unknown"

                #Status
                try:
                    status = issue.fields.status.name
                except:
                    status = "Unknown"

                #Reporter
                try:
                    reporter = issue.fields.reporter.displayName
                except:
                    reporter = "Unknown"



                issue_list = [
                    issue.key,
                    issue.id,
                    issue.fields.issuetype.name,
                    ttfr_complete,
                    ttfr_breach,
                    ttfr_time,
                    ttfr_remaining,
                    ttr_complete,
                    ttr_breach,
                    ttr_time,
                    ttr_remaining,
                    priority,
                    status,
                    issue.fields.summary,
                    created,
                    reporter,
                    assignee,
                    request_type,
                    product_type,
                    org,
                    resolution,
                    resolved,
                    updated
                ]

                #CMS-only fields
                if desk_id == "CMS":
                    try:
                        root_cause = issue.raw['fields']['customfield_11480'][0]['value']
                    except:
                        root_cause = ""

                    #Reason for breach
                    try:
                        reason_for_breach = issue.raw['fields']['customfield_11481'][0]['value']
                    except:
                        reason_for_breach = ""                  

                    #Reason for breach comment
                    try:
                        reason_for_breach_comment = issue.raw['fields']['customfield_11488'][0]['value']
                    except:
                        reason_for_breach_comment = ""

                    #Resolution()
                    try:
                        resolution_brackets = issue.raw['fields']['customfield_11447']['value']
                    except:
                        resolution_brackets = ""

                    #Job Type
                    try:
                        job_type = issue.raw['fields']['customfield_11471']['value']
                    except:
                        job_type = ""

                    #Department
                    try:
                        department = issue.raw['fields']['customfield_11470']
                    except:
                        department = ""

                    #Job Title
                    try:
                        job_title = issue.raw['fields']['customfield_11456']
                    except:
                        job_title = ""

                    # app_list = []
                    # try:
                    #     for i in range(len(issue.raw['fields']['customfield_11454'])):
                    #         app_item = issue.raw['fields']['customfield_11454'][i]['value']
                    #         app_list.append(app_item)
                    #     applications = ", ".join(app_list)
                    # except:
                    #     applications = ""

                    app_list = []

                    for i in range(10):
                        try:
                            app_list.append(issue.raw["fields"]["customfield_11454"][i]['value'])
                        except:
                            app_list.append("-")

                    #CSCC Action
                    try:
                        cscc_action = issue.raw['fields']['customfield_11435']['value']
                    except:
                        cscc_action = ""

                    #Report Page
                    try:
                        report_page = issue.raw['fields']['customfield_11458'][0]['value']
                    except:
                        report_page = ""

                    #Linked Issue
                    try:
                        linked_issue = issue.raw['fields']['issuelinks'][0]['inwardIssue']['key']
                        #linked_issue = issue.fields.issuelinks.inwardIssue.key
                    except:
                        linked_issue = ""
                    
                    #TTFR Orange
                    if issue.raw['fields']['customfield_11484'].get('ongoingCycle'):
                        ttfr_orange_completed = False
                        ttfr_orange_breach = issue.raw['fields']['customfield_11484']['ongoingCycle']['breached']
                        ttfr_orange_time = issue.raw['fields']['customfield_11484']['ongoingCycle']['elapsedTime']['millis'] / 1000 // 60
                        ttfr_orange_remaining = issue.raw['fields']['customfield_11484']['ongoingCycle']['remainingTime']['millis'] / 1000 // 60
                    elif issue.raw['fields']['customfield_11484'].get('completedCycles'):
                        ttfr_orange_completed = True
                        ttfr_orange_breach = issue.raw['fields']['customfield_11484']['completedCycles'][0]['breached']
                        ttfr_orange_time = issue.raw['fields']['customfield_11484']['completedCycles'][0]['elapsedTime']['millis'] / 1000 // 60
                        ttfr_orange_remaining = issue.raw['fields']['customfield_11484']['completedCycles'][0]['remainingTime']['millis'] / 1000 // 60
                    else:
                        ttfr_orange_completed = ""
                        ttfr_orange_breach = ""
                        ttfr_orange_time = ""
                        ttfr_orange_remaining = ""
                    

                    #TTR Orange
                    if issue.raw['fields']['customfield_11485'].get('ongoingCycle'):
                        ttr_orange_completed = False
                        ttr_orange_breach = issue.raw['fields']['customfield_11485']['ongoingCycle']['breached']
                        ttr_orange_time = issue.raw['fields']['customfield_11485']['ongoingCycle']['elapsedTime']['millis'] / 1000 // 60
                        ttr_orange_remaining = issue.raw['fields']['customfield_11485']['ongoingCycle']['remainingTime']['millis'] / 1000 // 60
                    elif issue.raw['fields']['customfield_11485'].get('completedCycles'):
                        ttr_orange_completed = True
                        ttr_orange_breach = issue.raw['fields']['customfield_11485']['completedCycles'][0]['breached']
                        ttr_orange_time = issue.raw['fields']['customfield_11485']['completedCycles'][0]['elapsedTime']['millis'] / 1000 // 60
                        ttr_orange_remaining = issue.raw['fields']['customfield_11485']['completedCycles'][0]['remainingTime']['millis'] / 1000 // 60
                    else:
                        ttr_orange_completed = ""
                        ttr_orange_breach = ""
                        ttr_orange_time = ""
                        ttr_orange_remaining = ""

                    #TTFR Premium
                    if issue.raw['fields']['customfield_11504'].get('ongoingCycle'):
                        ttfr_premium_completed = False
                        ttfr_premium_breach = issue.raw['fields']['customfield_11504']['ongoingCycle']['breached']
                        ttfr_premium_time = issue.raw['fields']['customfield_11504']['ongoingCycle']['elapsedTime']['millis'] / 1000 // 60
                        ttfr_premium_remaining = issue.raw['fields']['customfield_11504']['ongoingCycle']['remainingTime']['millis'] / 1000 // 60
                    elif issue.raw['fields']['customfield_11504'].get('completedCycles'):
                        ttfr_premium_completed = True
                        ttfr_premium_breach = issue.raw['fields']['customfield_11504']['completedCycles'][0]['breached']
                        ttfr_premium_time = issue.raw['fields']['customfield_11504']['completedCycles'][0]['elapsedTime']['millis'] / 1000 // 60
                        ttfr_premium_remaining = issue.raw['fields']['customfield_11504']['completedCycles'][0]['remainingTime']['millis'] / 1000 // 60
                    else:
                        ttfr_premium_completed = ""
                        ttfr_premium_breach = ""
                        ttfr_premium_time = ""
                        ttfr_premium_remaining = ""
                    

                    #TTR Premium
                    if issue.raw['fields']['customfield_11505'].get('ongoingCycle'):
                        ttr_premium_completed = False
                        ttr_premium_breach = issue.raw['fields']['customfield_11505']['ongoingCycle']['breached']
                        ttr_premium_time = issue.raw['fields']['customfield_11505']['ongoingCycle']['elapsedTime']['millis'] / 1000 // 60
                        ttr_premium_remaining = issue.raw['fields']['customfield_11505']['ongoingCycle']['remainingTime']['millis'] / 1000 // 60
                    elif issue.raw['fields']['customfield_11505'].get('completedCycles'):
                        ttr_premium_completed = True
                        ttr_premium_breach = issue.raw['fields']['customfield_11505']['completedCycles'][0]['breached']
                        ttr_premium_time = issue.raw['fields']['customfield_11505']['completedCycles'][0]['elapsedTime']['millis'] / 1000 // 60
                        ttr_premium_remaining = issue.raw['fields']['customfield_11505']['completedCycles'][0]['remainingTime']['millis'] / 1000 // 60
                    else:
                        ttr_premium_completed = ""
                        ttr_premium_breach = ""
                        ttr_premium_time = ""
                        ttr_premium_remaining = ""


                    #TTFR Premium Plus
                    if issue.raw['fields']['customfield_11510'].get('ongoingCycle'):
                        ttfr_premium_plus_completed = False
                        ttfr_premium_plus_breach = issue.raw['fields']['customfield_11510']['ongoingCycle']['breached']
                        ttfr_premium_plus_time = issue.raw['fields']['customfield_11510']['ongoingCycle']['elapsedTime']['millis'] / 1000 // 60
                        ttfr_premium_plus_remaining = issue.raw['fields']['customfield_11510']['ongoingCycle']['remainingTime']['millis'] / 1000 // 60
                    elif issue.raw['fields']['customfield_11510'].get('completedCycles'):
                        ttfr_premium_plus_completed = True
                        ttfr_premium_plus_breach = issue.raw['fields']['customfield_11510']['completedCycles'][0]['breached']
                        ttfr_premium_plus_time = issue.raw['fields']['customfield_11510']['completedCycles'][0]['elapsedTime']['millis'] / 1000 // 60
                        ttfr_premium_plus_remaining = issue.raw['fields']['customfield_11510']['completedCycles'][0]['remainingTime']['millis'] / 1000 // 60
                    else:
                        ttfr_premium_plus_completed = ""
                        ttfr_premium_plus_breach = ""
                        ttfr_premium_plus_time = ""
                        ttfr_premium_plus_remaining = ""
                    

                    #TTR Premium Plus
                    if issue.raw['fields']['customfield_11511'].get('ongoingCycle'):
                        ttr_premium_plus_completed = False
                        ttr_premium_plus_breach = issue.raw['fields']['customfield_11511']['ongoingCycle']['breached']
                        ttr_premium_plus_time = issue.raw['fields']['customfield_11511']['ongoingCycle']['elapsedTime']['millis'] / 1000 // 60
                        ttr_premium_plus_remaining = issue.raw['fields']['customfield_11511']['ongoingCycle']['remainingTime']['millis'] / 1000 // 60
                    elif issue.raw['fields']['customfield_11511'].get('completedCycles'):
                        ttr_premium_plus_completed = True
                        ttr_premium_plus_breach = issue.raw['fields']['customfield_11511']['completedCycles'][0]['breached']
                        ttr_premium_plus_time = issue.raw['fields']['customfield_11511']['completedCycles'][0]['elapsedTime']['millis'] / 1000 // 60
                        ttr_premium_plus_remaining = issue.raw['fields']['customfield_11511']['completedCycles'][0]['remainingTime']['millis'] / 1000 // 60
                    else:
                        ttr_premium_plus_completed = ""
                        ttr_premium_plus_breach = ""
                        ttr_premium_plus_time = ""
                        ttr_premium_plus_remaining = ""

                    #Weekly Status
                    try:
                        weekly_status = issue.raw['fields']['customfield_11453']
                    except:
                        weekly_status = ""

                    #Last public comment date
                    try:
                        last_public_comment_date = issue.raw['fields']['customfield_11444']
                    except:
                        last_public_comment_date = ""

                    #ERP Module
                    try:
                        erp_module = issue.raw['fields']['customfield_11302']['value']
                    except:
                        erp_module = ""

                    #Sales Agent Issue
                    try:
                        sales_agent_issue = issue.raw['fields']['customfield_11402']['value']
                    except:
                        sales_agent_issue = ""

                        #CRM Application
                    try:
                        crm_application = issue.raw['fields']['customfield_11441']['value']
                    except:
                        crm_application = ""


                    #PAYG Issue
                    try:
                        payg_issue = issue.raw['fields']['customfield_11438']['value']
                    except:
                        payg_issue = ""

                    #Payment Issue
                    try:
                        payment_issue = issue.raw['fields']['customfield_11437']['value']
                    except:
                        payment_issue = ""

                    
                    issue_list_cms = [root_cause, reason_for_breach, reason_for_breach_comment, resolution_brackets,
                                    job_type, department, job_title, app_list[0],app_list[1],app_list[2],app_list[3],
                                    app_list[4],app_list[5],app_list[6],app_list[7],app_list[9],app_list[9],
                                    cscc_action, report_page, linked_issue,
                                    ttfr_orange_completed,ttfr_orange_breach,ttfr_orange_time, ttfr_orange_remaining,
                                    ttr_orange_completed, ttr_orange_breach, ttr_orange_time,ttr_orange_remaining,
                                    ttfr_premium_completed,ttfr_premium_breach, ttfr_premium_time,ttfr_premium_remaining,
                                    ttr_premium_completed, ttr_premium_breach, ttr_premium_time,ttr_premium_remaining,
                                    ttfr_premium_plus_completed, ttfr_premium_plus_breach, ttfr_premium_plus_time,ttfr_premium_plus_remaining,
                                    ttr_premium_plus_completed, ttr_premium_plus_breach, ttr_premium_plus_time,ttr_premium_plus_remaining,
                                    weekly_status, last_public_comment_date, erp_module, sales_agent_issue,
                                    crm_application, payg_issue, payment_issue]

                    issue_list.extend(issue_list_cms)


                output_list.append(issue_list)


        #Write all metrics to Sharepoint
        fname = "../outputs/{}_jira_dump_{}.csv".format(run_date,desk_id)

        with open(fname, 'w', encoding='utf-8', newline='') as csvfile:
            
            print("Writing .csv file...")

            writer = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC)
            writer.writerow(header_list)
            writer.writerows(output_list)


        #Write to Sharepoint
            
        # file_content = read_file_as_binary(fname)

        # sp_path = 'teams/Engineering/Reporting/Service%20Desk%20Reporting/'

        # sp = Sharepoint(sharepoint_username,sharepoint_password, sharepoint_url)

        # sp.authorise_sharepoint(sharepoint_url)

        # sp.upload_file(file_content=file_content, file_name=fname, path=sp_path)
