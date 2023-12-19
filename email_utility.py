
import requests 
import re

# Microsoft Graph API endpoint
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'

# OAuth settings
ApplicationId= "8"
SecretId= "8"
SecretValue = "R"
tenant_id = "2"
# AUTHORITY = 'https://login.microsoftonline.com/'+tenant_id

def get_access_token():
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        'grant_type': 'client_credentials',
        'client_id': ApplicationId,
        'client_secret': SecretValue,
        'scope': 'https://graph.microsoft.com/.default'
    }
    response = requests.post(url, data=data)
    return response.json().get("access_token")

def send_email(recipient_email, subject, body):
    #url = 'https://graph.microsoft.com/v1.0/me/sendMail'
    url = "https://graph.microsoft.com/v1.0/users/chatbot@eevabitslab1.net/sendMail"
    access_token = get_access_token()
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    email_body = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": recipient_email
                    }
                }
            ]
        },
        "saveToSentItems": "true"
    }

    response = requests.post(url, headers=headers, json=email_body)
    if response.status_code == 202:
        print(f"Email sent to {recipient_email} successfully!")
    else:
        print(f"Failed to send email. Status code: {response.status_code}. Response text: {response.text}")
    return response

def send_reply(message_id, reply_content):
    """
    Send a reply to a message using MS Graph API.

    Returns:
    - dict: Dictionary containing response status and message.
    """
    
    token = get_access_token()

    # Step 2: Use the token to reply to the message
    headers = {
        'Authorization': 'Bearer {}'.format(token),
        'Content-Type': 'application/json',
    }
    reply_url = f'https://graph.microsoft.com/v1.0/users/chatbot@eevabitslab1.net/messages/{message_id}/reply'
    
    data = {
        "comment": reply_content
    }

    response = requests.post(reply_url, headers=headers, json=data)

    # Check response
    if response.status_code == 202:
        return {"status": "success", "message": "Reply successfully created!"}
    else:
        return {"status": "error", "message": f"Failed to create the reply. Status code: {response.status_code}", "details": response.text}

def get_unread_emails(subject):
    # API endpoint to get messages
    endpoint = "https://graph.microsoft.com/v1.0/users/chatbot@eevabitslab1.net/messages"
    access_token = get_access_token()
    # Filter for unread messages
    filters = "$filter=isRead eq false"
    # Select only ID and subject 
    # select = "$select=id,subject, "
    # select = "$select=id,subject,from,bodyPreview"
    # Construct full URL
    full_url = endpoint + "?" + filters 

    # Make API request
    response = requests.get(full_url,  headers={'Authorization': 'Bearer ' + access_token})

    # Get emails from response
    conversation_ids = []
    if response.status_code == 200:
        emails = response.json()['value']
        for email in emails:
            # print(email)
            if email['subject'] is not None and subject in email['subject'] and email['conversationId'] not in conversation_ids:
                conversation_ids.append(email['conversationId'])
    return conversation_ids

def check_folders_validation_message_sent(email_id, subject, folders):
    """
    Get emails by its ID using Microsoft Graph API.
    """
    # Define the URL for fetching the email by ID
    endpoint = f'https://graph.microsoft.com/v1.0/users/chatbot@eevabitslab1.net/messages?$search="participants:{email_id}"'
    access_token = get_access_token()
    headers = {
        "Authorization": "Bearer " + access_token
    }
    response = requests.get(endpoint, headers=headers)
    messages = response.json()["value"]
    email_messages = ''
    for message in messages:
        if subject in message["subject"]:
            body = extract_text_from_html(message['body']['content'])
            email_messages += body
    checked_folders = {}
    for folder in folders:
        if email_messages != '' and folder in email_messages:
            checked_folders[folder] = True
        else:
            checked_folders[folder] = False
            
    return checked_folders

from bs4 import BeautifulSoup

message = ""

# Function to recursively get text from elements
def get_text(element):
    try:
        global message
        div = element.find('body')
        if div is not None:
            for text in div.stripped_strings:
                message += text+'\n'

        if message.strip() != '':
            message = message.split('From:')[0]
    except Exception as ex:
        print(ex)
        import os, sys
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)

def extract_text_from_html(html_string):
    global message
    # Create a BeautifulSoup object
    soup = BeautifulSoup(html_string, 'html.parser')
    # Get all the text from the HTML
    message = ""
    # Start the recursive process
    get_text(soup)
    # text = soup.get_text()
    return message

def get_thread_id(conversation_id):
    """
    Get the thread ID of a conversation using Microsoft Graph API.

    Args:
    - conversation_id (str): The ID of the conversation.
    - access_token (str): The access token for authentication.

    Returns:
    - str: The thread ID of the conversation.
    """
    
    # Define the URL for fetching messages in the conversation
    messages_url = f'https://graph.microsoft.com/v1.0/users/chatbot@eevabitslab1.net/messages?$filter=conversationId eq \'{conversation_id}\''
    access_token = get_access_token()
    # Make the API request
    response = requests.get(
        messages_url,
        headers={'Authorization': f'Bearer {access_token}'}
    )

    if response.status_code == 200:
        messages_data = response.json().get('value', [])
        if messages_data:
            return messages_data[0]['conversationId'], messages_data[0]['conversationThreadId']
        else:
            print(f"No messages found in the conversation.")
            return None, None
    else:
        print(f"Failed to fetch messages. Status code: {response.status_code}")
        return None, None

def get_thread(conversation_id=None):
    from urllib.parse import quote
    messages_url = f"https://graph.microsoft.com/v1.0/users/chatbot@eevabitslab1.net/messages?$filter=conversationId eq '{conversation_id}'"
    access_token = get_access_token()
    # conversation_messages_id, thread_id = get_thread_id(conversation_id)
    # print(conversation_messages_id)
    # print(thread_id)

    # print(conversation_id)

    response = requests.get(messages_url, 
                            headers={'Authorization': 'Bearer ' + access_token})
    messages = response.json().get('value', [])
    conversation_text = ''
    thread_messages = []
    for message in messages[-1::-1]:
        # pass
        # print( message)
        # print('='*100)
        if  message['conversationId']  == conversation_id:
            thread_messages.append(message)
            _from = 'From: '+message['sender']['emailAddress']['address']
            _to = 'To: '+message['toRecipients'][0]['emailAddress']['address']

            body = message['body']['content'] #if message['body']['contentType'] == 'text' else message['bodyPreview']
            # print(body)
            # print('='*200)
            if '<html>' in body:
                body = extract_text_from_html(body)
            # print(body)
            if 'wrote:' in body:
                body = re.sub(r'(?s)\n\s*On.*wrote:.*$', '', body, re.DOTALL)
                body = body.split('From:')[0]
                body = re.sub(r'On\s+[A-Za-z]{3},\s*[A-Za-z]{3}\s*\d{1,2},\s*\d{4}.*$', '', body, re.DOTALL)
            
            message_text = '\n\n'+_from+'\n'+_to+'\n'+'\n'+ body
            conversation_text = message_text + '\n\n'+ conversation_text
    print(conversation_text)
    owner_email = ''
    if thread_messages[0]['sender']['emailAddress']['address'] !=  'chatbot@eevabitslab1.net':
        owner_email =  thread_messages[0]['sender']['emailAddress']['address'] 
    else:
        owner_email =  thread_messages[0]['toRecipients'][0]['emailAddress']['address']
    return conversation_text, thread_messages[0]['id'], owner_email

def mark_as_read(conversation_id=None):
    messages_url = "https://graph.microsoft.com/v1.0/users/chatbot@eevabitslab1.net/messages"
    access_token = get_access_token()
    response = requests.get(messages_url, 
                            headers={'Authorization': 'Bearer ' + access_token})

    messages = response.json()['value']
    unread_msgs = [m for m in messages if  m['conversationId'] == conversation_id and not m['isRead']]

    for message in unread_msgs:
        try:
            endpoint = f"https://graph.microsoft.com/v1.0/users/chatbot@eevabitslab1.net/messages/{message['id']}"
            headers = {
                            'Authorization': f'Bearer {access_token}',
                            'Content-Type': 'application/json'
                    }
            # The data payload specifying that the email should be marked as read
            data = {
                'isRead': True
            }

            response = requests.patch(endpoint, headers=headers, json=data)

            if response.status_code == 200:
                print("Email marked as read successfully!")
            else:
                print("Failed to mark email as read. Status code:", response.status_code)
        except Exception as ex:
            print(ex)
            continue

def filter_and_read_emails_by_email_address_and_subject(email_address: str, subject: str) -> list[dict]:
    """Filters and reads emails by email address and subject from the Microsoft Graph API.

    Args:
    email_address: The email address to filter by.
    subject: The subject to filter by.

    Returns:
    A list of dictionaries containing the email data, or an empty list if no emails were found.
    """
    try:
        # Get the access token for the Microsoft Graph API.
        # Replace `CLIENT_ID` and `CLIENT_SECRET` with your own values.
        access_token = get_access_token()

        # Construct the request URL.
        #messages_url = f"https://graph.microsoft.com/v1.0/users/chatbot@eevabitslab1.net/messages?$filter=conversationId eq '{conversation_id}'"
        url = f"https://graph.microsoft.com/v1.0/users/chatbot@eevabitslab1.net/messages?$toRecipients/any(r: r/emailAddress/address eq '{email_address}') and contains(subject,'{subject}')"

        # Make the request.
        headers = {
        "Authorization": f"Bearer {access_token}"
        }
        response = requests.get(url, headers=headers)

        # Check the response status code.
        if response.status_code != 200:
            return []

        # Get the email data from the response.
        email_data = response.json()
        # Return the email data.
        message_sent = False
        for message in email_data["value"]:
            if len(message['toRecipients']) > 0 and  message['toRecipients'][0]['emailAddress']['address'] == email_address and subject in message['subject']:
                message_sent = True
                break
        return message_sent
    except Exception as ex:
        print(ex)
        import sys, os
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        return False
