from urllib import response
import msal
import requests
import json
import numpy as np
import pandas as pd
import webbrowser
import msal
import os
import re
from bs4 import BeautifulSoup
import cohere
import streamlit as st
import seaborn as sns
import altair as alt

# ------------------------------------
# ------- Graph API Configuration ----
# ------------------------------------
APPLICATION_ID = 'ea810a22-d8f7-4399-8d28-8c817e7ad762'
# DIRECTORY_ID = '6601e215-7260-414a-850a-d0cde0136c40'
# CLIENT_SECRET = 'RTz8Q~efpfq8OfYIvPiAiSjEKy.0MLN8re-BybrX'

authority_url = 'https://login.microsoftonline.com/consumers/'
base_url = 'https://graph.microsoft.com/v1.0/'

SCOPES = [
    'Mail.Read',
    'Mail.ReadWrite',
    'Mail.ReadBasic',
]

# Create a function to generate an access token
def generate_access_token(app_id, scopes):
    access_token_cache = msal.SerializableTokenCache()

    # Read the token file
    if os.path.exists('api_token_access.json'):
        access_token_cache.deserialize(open('api_token_access.json', 'r').read())
    
    client = msal.PublicClientApplication(app_id, token_cache=access_token_cache)

    accounts = client.get_accounts()
    if accounts:
        token_response = client.acquire_token_silent(scopes, account=accounts[0])
    else:
        flow = client.initiate_device_flow(scopes=scopes)
        app_code = flow['message']
        print(app_code)
        st.write(app_code)
        webbrowser.open(flow['verification_uri'], new=True)

        token_response = client.acquire_token_by_device_flow(flow)

        with open('api_token_access.json', 'w') as f:
            f.write(access_token_cache.serialize())

    return token_response['access_token']

# Function for generating emails
def all_mail(access_token, number_of_mails):
    endpoint = base_url + 'me/messages?$top=' + str(number_of_mails)
    headers = {
        'Authorization': 'Bearer ' + access_token
    }
    response = requests.get(endpoint, headers=headers)
    return response.json()

def get_mail_from_me(access_token, fromEmailAddress, number_of_mails):
    endpoint = base_url + 'me/messages?$filter=from/emailAddress/address%20eq%20%27' + fromEmailAddress + '%27&$top=' + str(number_of_mails)
    headers = {
        'Authorization': 'Bearer ' + access_token
    }
    response = requests.get(endpoint, headers=headers)
    return response.json()

def get_email_sent_to_me(access_token, toEmailAddress, number_of_mails):
    endpoint = base_url + 'me/messages?$filter=toRecipients/emailAddress/address%20eq%20%27' + toEmailAddress + '%27&$top=' + str(number_of_mails)
    headers = {}
    pass

# Function for getting emails where the inferenceClassification is set to 'focused'
def focused_mail(access_token, number_of_mails):
    endpoint = base_url + 'me/messages?$filter=inferenceClassification%20eq%20%27focused%27&$top=' + str(number_of_mails)
    headers = {
        'Authorization': 'Bearer ' + access_token
    }
    response = requests.get(endpoint, headers=headers)

    # If the response is not 200, then raise an exception
    if response.status_code != 200:
        raise Exception(response.text)

    return response.json()

# Get the body of the email
def get_body_text(response):
    """Extract the raw HTML content from the body of an email."""
    body_text = None
    try:
        body = response['body']
        body_text = body['content']
    except (KeyError, IndexError):
        pass
    
    return body_text

# Get the text from the email
def get_text_from_html(html):
    """Extract the text from an HTML document."""
    soup = BeautifulSoup(html, "html.parser")
    # Get the text from all the direct children of the body tag.
    text = soup.body.get_text(separator="\n")
    # Get rid of all script and style elements.
    for script in soup(["script", "style"]):
        script.extract()
    # Get the text from all the direct children of the body tag including span and div tags.
    text = soup.body.get_text(separator="\n")
    # Collapse multiple newlines
    return re.sub(r"[\r\n]+", "\n", text)

# Parse the email object and retain only the relevant information in a dictionary
def parse_email(response):
    """Parse the email object and retain only the relevant information in a dictionary."""
    email = {}
    try:
        email['id'] = response['id']
        email['subject'] = response['subject']
        email['receivedDateTime'] = response['receivedDateTime']
        email['from'] = response['from']['emailAddress']['address']
        email['to'] = response['toRecipients'][0]['emailAddress']['address']
        email['inferencedClassification'] = response['inferenceClassification']
        email['body'] = get_text_from_html(get_body_text(response))

    except (KeyError, IndexError):
        print("Error parsing email")
        pass

    return email

def structure_prompt(email):
    """Structure the prompt for the API call."""
    prompt = "Subject: " + email['subject'] + "\n"
    prompt += "From: " + email['from'] + "\n"
    prompt += "To: " + email['to'] + "\n"
    prompt += "Date: " + email['receivedDateTime'] + "\n"
    prompt += "Content: " + email['body'] + "\n"
    prompt += "\n###\n\n"

    return prompt

def send_email(subject, body, toEmailAddress):

    access_token = generate_access_token(APPLICATION_ID, SCOPES)

    endpoint = base_url + 'me/sendMail'
    headers = {
        'Authorization': 'Bearer ' + access_token,
    }

    request_body = {
    "message": {
        "subject": subject,
        "body": {
            "contentType": "Text",
            "content": body
        },
        "toRecipients": [
            {
                "emailAddress": {
                    "address": toEmailAddress
                }
            }
        ]
    }
    }

    response = requests.post(endpoint, headers=headers, json=request_body)
    print(response.text)

    if response.status_code != 202:
        raise Exception(response.text)




if __name__ == '__main__':
    # Generate an access token
    access_token = generate_access_token(APPLICATION_ID, SCOPES)

    # # Define the endpoint URL for getting the messages 
    # url = "https://graph.microsoft.com/v1.0/me/messages"

    # # Define the query parameters
    # params = {
    #     "$filter": "toRecipients/emailAddress/address eq 'nikhil.sehgal@vastmindz.com'",
    #     "$select" : "subject,receivedDateTime"
    # }

    # headers = {
    #     'Authorization': 'Bearer ' + access_token,
    # }

    # # Make a GET request to the endpoint
    # response = requests.get(url, params=params, headers=headers)

    # # Parse the response as JSON
    # data = json.loads(response.text)

    # print(data)

    # # Extract the messages from the response
    # messages = data["value"]

    # # Iterate through each message and retrieve the response
    # for message in messages:
    #     message_id = message["id"]
    #     received_time = message["receivedDateTime"]
    #     subject = message["subject"]
    #     message_content = message["body"]["content"]
    #     response_url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/reply"
    #     response = requests.get(response_url, headers=headers)
    #     response_data = json.loads(response.text)
    #     response_body = response_data["body"]["content"]

    #     # Write the message content and the response body to email.txt file
    #     with open('email.txt', 'a') as f:
    #         f.write(message_content + "\n\n")
    #         f.write(response_body + "\n\n")
    #         f.write("##############################################\n\n")



    # # Get the emails
    # emails = all_mail(access_token, 10)

    # structured_emails = []
    
    # for email in emails['value']:
    #     try:
    #         # Parse the email object and retain only the relevant information in a dictionary
    #         parsed_email = parse_email(email)
    #         # Structure the prompt for the API call
    #         prompt = structure_prompt(parsed_email)
    #         # Append the prompt to the list of prompts
    #         structured_emails.append(prompt)
    #     except:
    #         continue
        
    # # Write the prompt to the emails.txt file and separate each email with a line of hashes
    # with open('emails.txt', 'w') as f:
    #     f.write('\n\n----------------------------------\n\n'.join(structured_emails))