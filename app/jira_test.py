from jira import JIRA
import json
import sys
from datetime import datetime as dt

if __name__ == "__main__":

    jira = JIRA('https://bboxxltd.atlassian.net', basic_auth=('p.homer@bboxx.co.uk', 'S0trwuh#1'))

    issue = jira.issue('CMS-4221')


    print(json.dumps(issue.raw, sort_keys=True, indent=4, separators=(',', ': ')))

    