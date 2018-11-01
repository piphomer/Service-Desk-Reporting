#!/usr/bin/env python
#-*- coding: utf-8 -*-

from jira import JIRA
import csv
import sys
import os
import requests
from datetime import datetime as dt
import unicodedata
from unidecode import unidecode


from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.file_creation_information import FileCreationInformation
from office365.runtime.utilities.request_options import RequestOptions

#List of service desks to iterate through
service_desk_list = ['CMS','PS', 'ES']

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
        print('Uploading File "%s" to sharepoint folder "%s" ...' % (info.url, folder_object))
        print(full_url)
        response = requests.post(
            url=full_url, data=file_content, headers=options.headers, auth=options.auth,
        )
        if response.status_code not in [200,201]:
            raise Exception(response.text, response.status_code)
        print('Done.')



    def delete_file(self, path):
        print "Deleting file"
        split_path = path.rsplit('/', 1)
        folder_path, filename = split_path[0], split_path[1]
        folder = self.get_folder_by_server_relative_url(folder_path)
        self.ctx.load(folder)
        self.ctx.execute_query()

        files = folder.files
        self.ctx.load(files)
        self.ctx.execute_query()

        print files

        for cur_file in files:
            print cur_file.properties["Name"]
            if cur_file.properties["Name"] == filename:
                full_url = "{0}/$value".format(cur_file.url)
                options = RequestOptions(full_url)
                self.ctx.authenticate_request(options)
                self.ctx.ensure_form_digest(options)
                response = requests.delete(url=cur_file.url, headers=options.headers, auth=options.auth)

                print response.status_code

                if response.status_code == 404:
                    raise SharepointFileNotFound()

                elif response.status_code != 200:
                    raise SharepointDeleteFailed()

        return


def read_file_as_binary(path):
    with open(path, 'rb') as content_file:
        file_content = content_file.read()
    return file_content


####################################################################################################


if __name__ == "__main__":

    run_date = dt.strftime(dt.now(), '%y%m%d')

    print run_date

    # Connect to BBOXX Jira server
    jira = JIRA('https://bboxxltd.atlassian.net', basic_auth=(jira_username, jira_password))

    #Create connection to Sharepoint

    sp = Sharepoint(sharepoint_username,sharepoint_password, sharepoint_url)

    sp.authorise_sharepoint(sharepoint_url)
    
    #Move any files in the Active folder to the Archive folder
  
    #Coming soon...


    #Get the tickets

    for desk_id in service_desk_list:

        search_string = 'project = ' + desk_id

        print search_string

        # Retrieve all tickets so we can count them.
        # Note: we can only query 50 at a time so we will need to paginate in a later step
        all_tix = jira.search_issues(search_string)

        tix_count = all_tix.total # Count the tickets via the .total attribute

        page_qty = tix_count / 100 + 1# Calculate how many pages of tickets there are

        #page_qty = 1

        print "Total number of tickets: ",tix_count
        print "Number of pages: ", page_qty

        output_list = []
        output_list_debug = []

        # Loop through the number of pages we need to gather all issues
        for page in range(page_qty):

            issue_list = []
            issue_list_debug = []

            #print "Page: ", page + 1
            print "\r" + "Page ", page + 1, " of ", page_qty,

            starting_issue = page * 100

            tix = jira.search_issues(search_string, startAt = starting_issue, maxResults= 100)

            for issue in tix:
                issue_priority = issue.fields.priority.name

                # print issue.key, "..."

                #time to first response
                try:
                    #Check if first response has occurred
                    ttfr = issue.raw['fields']['customfield_10806']["ongoingCycle"]["remainingTime"]["friendly"]
                    #But just return that we are awaiting it, not the time left till breach
                    ttfr = 'Awaiting first response'
                    ttfr_mins = ""
                    ttfr_status = "Ongoing"
                    ttfr_time = ""
                except:
                    # If no first response has occurred yet, then "ongoingCycle" above did not exist
                    try:
                        ttfr = issue.raw['fields']['customfield_10806']["completedCycles"][0]["remainingTime"]["friendly"]
                        # cf10806 returns time left till SLA breach
                        # So, subtract it from SLA duration (in mins) to get the actual elapsed time.
                        if desk_id == "CMS":
                            if issue_priority == "High":
                                sla_mins = 240
                            elif issue_priority == "Medium":
                                sla_mins = 480
                            else:
                                sla_mins = 1440
                        else:
                            sla_mins = 480
                        ttfr_mins = sla_mins - issue.raw['fields']['customfield_10806']['completedCycles'][0]['remainingTime']['millis']/1000/60
                        ttfr_mins = str(ttfr_mins)
                        ttfr_status = str(issue.raw['fields']['customfield_10806']["completedCycles"][0]["breached"])
                        ttfr_time = issue.raw['fields']['customfield_10806']["completedCycles"][0]["stopTime"]["jira"]
                    except:
                        #Some SLAs seem to get deleted. If so return "??"
                        ttfr = "??"
                        ttfr_status = "??"
                        ttfr_time = "??"
                        ttfr_mins = "??"

                
                #time to resolution
                if desk_id == "CMS" or desk_id == 'ES':
                    if issue_priority == "High":
                        sla_mins = 2400
                    elif issue_priority == "Medium":
                        sla_mins = 3600
                    else:
                        sla_mins = 4800
                    try:
                        #Let's see if there is any remaining time information by trying it (but not using it)
                        ttr = issue.raw['fields']['customfield_10805']["ongoingCycle"]["remainingTime"]["friendly"]
                        #If there was, just say we are awaiting resolution
                        ttr = "Awaiting resolution" #Just overwrite the value we got through trying
                        ttr_mins = ""
                        ttr_status = "Ongoing"
                    except:
                        try:
                            ttr = issue.raw['fields']['customfield_10805']["completedCycles"][0]["remainingTime"]["friendly"]
                            ttr_status = str(issue.raw['fields']['customfield_10805']["completedCycles"][0]["breached"])
                            ttr_mins = sla_mins - issue.raw['fields']['customfield_10805']['completedCycles'][0]['remainingTime']['millis']/1000/60
                            ttr_mins = str(ttr_mins)
                        except: #Go down this branch if there is no TTR information at all (ticket was transferred from PS, for example)
                            ttr = "??"
                            ttr_status = "??"
                            ttr_mins = "??"
                else:
                    #There is no SLA for TTR on PS desk so the custom field does not get populated with anything useful
                    #So we will need to just look at the difference between raised date and resolved date            
                    if issue.fields.resolutiondate != None:
                        ttr = ""
                        ttr_status = "n/a"
                        ttr_mins = dt.strptime(issue.fields.resolutiondate[:16], '%Y-%m-%dT%H:%M') - dt.strptime(issue.fields.created[:16], '%Y-%m-%dT%H:%M')
                        ttr_mins = str(ttr_mins.seconds/60)
                    else:    
                        ttr = ""
                        ttr_status = ""
                        ttr_mins = ""

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
                    ttfr,
                    ttfr_mins,
                    ttfr_status,
                    responded,
                    ttr,
                    ttr_mins,
                    ttr_status,
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


                #Encode unicode fields to bytes
                for i, item in enumerate(issue_list):
                    
                    if (type(issue_list[i]) != 'str' and type(issue_list[i]) != None):
                    #     issue_list[i] = unicodedata.normalize('NFKD',item).encode('utf-8', errors='replace')
                        try:
                            issue_list[i] = issue_list[i].decode('utf-8', errors='replace')
                        except:
                            issue_list[i] = "Unicode decode error... sorry!!"


                output_list.append(issue_list)


        #Write all metrics to Sharepoint
        fname = "{}_jira_dump_{}.csv".format(run_date,desk_id)

        with open(fname, 'wb') as csvfile:
            
            print "Writing .csv file..."
            writer = csv.writer(csvfile, quoting=csv.QUOTE_NONNUMERIC)
            writer.writerow([
                "Ticket Key",
                "Ticket ID",
                "Issue Type",
                "Time to First Response",
                "TTFR (Minutes)",
                "TTFR Breached",
                "Time of First Response",
                "Time to Resolution",
                "TTR (Minutes)",
                "TTR Breached",
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
            ])
            writer.writerows(output_list)


            #Write to Sharepoint
            
            # file_content = read_file_as_binary(fname)

            # sp_path = 'teams/Engineering/Reporting/Service%20Desk%20Reporting/'

            # sp.upload_file(file_content=file_content, file_name=fname, path=sp_path)
