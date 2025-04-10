
"""
Hubspot Email Analytics Extractor - March 2025

This script connects to the HubSpot API to retrieve email campaign data,
including subject lines, publish dates, open/deliver rates, and target
lists. It outputs report rows for further analysis.

ssoStatus(): whether an email was sent to single sign-on users or not
getEmails(): returns the id and subject of emails sent since Nov 1, 2024
createRow(): creates a spreadsheet row based on the details of an email (stats, distributed to, etc)
getAllRows(): grabs emails from getEmails(), createRow() for each email, and returns all the rows in one list

"""



import requests
from datetime import datetime, timezone
from secrets import HUBSPOT_KEY as KEY

hubspot_url = "https://api.hubapi.com/marketing/v3/emails"

headers = {
    "accept": "application/json",
    "authorization": f"Bearer {KEY}"
}





def ssoStatus(num):

    """
    Whether the email was sent to single sign-on users or not.
    
    """
    
    if num == ['6922']: return "SSO"
    
    elif num == ['6927']: return "noSSO"
    
    else: return ""




    
def getEmails():
    
    """
    Fetches all HubSpot marketing emails created after November 1st, 2024.
    Returns a list of dictionaries with the email 'id' and 'subject'.
    
    """
    
    created_after = int(datetime(2024, 11, 1).timestamp() * 1000)

    general_params = {
        "createdAfter": created_after
    }


    emails = []
    has_more = True
    after = None

    # Pagination: keep fetching until there are no more pages
    while has_more:
        if after:
            general_params["after"] = after

        # Make the request to the API
        response = requests.get(hubspot_url, headers=headers, params=general_params)
        data = response.json()

        # Add the emails from the current page to the list
        emails.extend([
            {"id": email["id"], "subject": email.get("subject", "No subject")}
            for email in data.get("results", [])
        ])

        # Check if there are more results & update 'after'
        has_more = data.get("paging", {}).get("next", {}).get("after", None) is not None
        after = data.get("paging", {}).get("next", {}).get("after")

        
    return emails






def createRow(details):

    """
    Creates a spreadsheet row for each list a particular email was sent to:
    email_id, heading, date_str, time_str, total_delivers, total_opens, listName, ssoStatus
    """

    rows = []
    
    # General Info
    email_id = details['id']
    isPublished = details['isPublished']

    if isPublished:
        datetime_str = details['publishDate']
        datetime_obj = datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%SZ")
        date_str = datetime_obj.date().isoformat()
        time_str = datetime_obj.time().isoformat()
        
        subject_line = details['subject']


        # Open Rate
        try:
            stats = details['stats']['counters']
            total_delivers = stats['delivered']
            total_opens = stats['open']
        except:
            print(f"Error getting engagement states for {heading}")


        # Distribution
        try:
            send_to = details['to']['contactIlsLists']['include']
            dont_send_to = details['to']['contactIlsLists']['exclude']
            ssoStatus = ssoStatus(dont_send_to)
        except:
            print(f"Error getting distribution for {heading}")
            return

        lists = []

        for list_id in send_to:
            list_url = f"https://api.hubapi.com/crm/v3/lists/{list_id}"
            response = requests.get(list_url, headers=headers)
            data = response.json()
            name = data['list']['name']
            if name != "Staff":
                lists.append(name)
                    
        for listName in lists:
            if "Distro List" in listName:
                row = [email_id, subject_line, date_str, time_str, "", "", total_delivers, total_opens, listName, ssoStatus]
                rows.append(row)

        return rows




def getAllRows():

    detail_params = {
        "includeStats": True
    }
        
    emails = getEmails()
    all_rows = []
    
    for email in emails:
        email_id = email['id']
        detail_url = f"{url}/{email_id}"
        data = requests.get(detail_url, headers=headers, params=detail_params)
        details = data.json()

        rows = createRow(details)

        for row in rows: all_rows.append(row)

    return all_rows


        

