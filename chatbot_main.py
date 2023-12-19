
import openai
import os
import re
import time
import pyodbc

import smtplib
import imaplib
import email
# import langchain
import json
from email.message import EmailMessage
import json
import requests
import email_utility as util
import openAI_utility as ai_util
import db_utility as db_util
# Set up IMAP for reading emails 
OUTLOOK_SMTP = "smtp.office365.com"
OUTLOOK_PORT = 465
OUTLOOK_EMAIL = "chatbot@******.net"
OUTLOOK_PASSWORD = "******"

smtp_server = "smtp.office365.com"
port = 465 #if use_ssl else 587


def get_initial_email_template(owner_name, owner_email, resource_ids):
    """
    Writes an email to the owner listing the resource in the email. Asking the owner if they own the resources and send it to the owner.
    """

    subject = 'Data Migration to Microsoft Teams Task: Managed Folder(s) Validation'
    email_body = f"""Dear {owner_name},
    I'm the Contoso IT Chatbot assisting with the ongoing data migration project to Microsoft Teams. Based on our data analysis, we've identified the following Managed Folders as primarily utilized by the Design team under your leadership:
    
    Folders:
    {resource_ids}
    
    Could you please review the folders and confirm if they should remain under your ownership? If any of the folders should be overseen by a different individual or team, kindly specify which and recommend an appropriate owner. Your guidance is instrumental in ensuring a seamless migration process.
    
    Warm regards,
    Contoso IT Chatbot
    """
    template = {
                'Email_ID': owner_email,
                'Owner' : owner_name,
                'Email_Subject': subject,
                'Email_Body': email_body
                }
    
    return template

def get_name_confirmation_email_tempalte(owners, reciever_name):
    """
    Writes an email to the owner listing the resource in the email. Asking the owner to confirm the name of the suggested owner.
    """
    suggested_owners = ''
    for i, owner in enumerate(owners):
        suggested_owners += '\n'+str(i+1)+': '+owner[0]+', '+owner[1]+', '+owner[2]

    # subject = 'Data Migration to Microsoft Teams Task: Managed Folder(s) Validation'
    
    email_body = f"""
    Dear {reciever_name},
    \n\rCould you help us determine which one you are referring to?

    \n\r{suggested_owners}. 
     
    \n\rWarm regards,
    \n\r\rContoso IT Chatbot
    """
    
    return email_body

def get_final_confirmation_email_tempalte(reciever_name, accepted_folders, rejected_folders ):
    """
    Writes an email to the owner listing the resource in the email. Asking the owner to reconfirm accepted and rejected folders.
    """

    accepted_list = ''
    rejected_list = ''
    try:
        for i, folder in enumerate(accepted_folders):
            accepted_list += '\n'+str(i+1)+'. '+folder['folder']

        for i, folder in enumerate(rejected_folders):
            print(' ---- ',folder['folder'], folder['names'][0])
            rejected_list += '\n'+str(i+1)+'. '+folder['folder']+' '+str(folder['names'][0])+''

        email_body = f"""Dear {reciever_name},

        Thank you for your response. As confirmation, you are accepting Data Ownership for the following Managed Folders:\n
        {accepted_list}.\n
        
        You have rejected ownership for the following Managed Folders:

        {rejected_list}.\n

        Please confirm this information is correct. \n
        
        Warm regards,
        \n
        Contoso IT Chatbot
        """
        return email_body
    except Exception as ex:
        print(ex)
        import sys
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)


if __name__ == '__main__':
    try:
        #================================================================================================
        #               Look for invalid resources
        #================================================================================================
        util.filter_and_read_emails_by_email_address_and_subject('abdulwali.nuces@gmail.com', 'Data Migration to Microsoft Teams Task')
        resources = db_util.get_invalid_resources()
        if resources:
            for owner_sid, resources_list in resources.items():
                try:
                    # Check if the file exists
                    owner_details = db_util.get_owner_details(owner_sid)
                    # Is Email already sent
                    folders_info = util.check_folders_validation_message_sent(owner_details[2], 'Data Migration to Microsoft Teams Task', resources_list)
                    resources_for_email = ''
                    for r, res in enumerate(folders_info):
                        if folders_info[res] is False:
                            resources_for_email += '\n'+str(r+1)+': '+res
                    
                    if owner_details is not None and resources_for_email.strip() != '' and owner_details[1].split()[0] == 'Brian':
                        # Check if email already not sent then proceed
                        result = db_util.managedata(owner_details[2])
                        print(owner_details[2],'\n' ,result)
                        email_sent = util.filter_and_read_emails_by_email_address_and_subject(owner_details[2], 'Data Migration to Microsoft Teams Task')
                        if not email_sent:
                            template = get_initial_email_template(owner_details[1].split()[0], owner_details[2],  resources_for_email)
                            #print('Email Template:\n',template)
                            #print('='*100)
                            util.send_email(template['Email_ID'], template['Email_Subject'],template['Email_Body'])
                except Exception as ex:
                    print(ex)
                    continue
        

        #================================================================================================
        #    Check for unread owner email responses with subject of data migration validation 
        #================================================================================================
        conversation_ids  = util.get_unread_emails('Data Migration to Microsoft Teams Task: Managed Folder(s) Validation')
        print('conversation_ids -- ',conversation_ids) 
        if conversation_ids:
            for conv_id in conversation_ids:
                print('conv_id -- ',conv_id)
                try: 
                    conversation_text = ''
                    conversation_text, message_id, email_id = util.get_thread(conv_id)
                    
                    matches = re.findall(r"\\\\[A-Za-z0-9\\]+\s", conversation_text, re.DOTALL)
                    folders_ids = {}

                    for mat in matches:
                        folder = mat.strip().strip('\n')
                        if folder not in folders_ids.keys():
                            folders_ids[folder]= db_util.get_resource_id(folder)

                    for folder, id_ in folders_ids.items():
                        conversation_text = conversation_text.replace(folder, 'folder-r'+str(id_))
                    
                    print('='*100)
                    # print('conversation_text --- ',conversation_text)
                    result = db_util.get_owner_details_by_email(email_id)
                    owner_name = ''
                    if result:
                        owner_name = result[1].split()[0]

                    # # print(conversation_text)
                    util.mark_as_read(conv_id)
                    #print(conversation_text)
                    prompt = ai_util.get_prompt(conversation_text)
                    if len(prompt.split()) < 4000:
                        # print(prompt)
                        response_text = ai_util.process_prompt(prompt)
                        # print('response_text -- ', response_text)
                        matched = re.search(r"\[.*\]", response_text, re.DOTALL)
                        json_content_greedy = matched.group() if matched else None
                        
                        owners = json.loads(json_content_greedy)
                        for owner in owners:
                            print(owner) 
                        if owners and 'Could you help us determine which one you are referring to?' not in conversation_text:
                            all_owners_okay = True
                            for owner in owners:
                                if 'email' not in owner:
                                    owner['email'] = 'NA'
                                
                                newowner = db_util.check_owners_exists(owner)
                                if newowner['status'] == 'reconfirm':
                                    all_owners_okay = False
                                    
                                    template = get_name_confirmation_email_tempalte(newowner['names'], owner_name)
                                    response = util.send_reply(message_id, template)
                        else:
                            accepted_folders = []
                            rejected_folders = []
                            for owner in owners:
                                if 'email' not in owner:
                                    owner['email'] = 'NA'
                                newowner = db_util.check_owners_exists(owner)
                                if newowner['status'] == 'valid':
                                    accepted_folders.append(newowner)
                                elif newowner['status'] == 'reconfirm':
                                    rejected_folders.append(newowner)

                            print('accepted_folder -- ',accepted_folders)
                            print('rejected_folders -- ',rejected_folders)
                            print('='*100)
                            prompt = ai_util.get_accepted_resources_verifiactio_prompt(conversation_text)
                            # print(prompt)
                            # print('='*100)
                            response_text = ai_util.process_prompt(prompt)
                            # print(response_text)
                            matched = re.search(r"\{.*\}", response_text, re.DOTALL)
                            json_content_greedy = matched.group() if matched else None
                            status = json.loads(json_content_greedy)
                            print('status -- ',status)

                            if "As confirmation, you are accepting Data Ownership for the following Managed Folders:" not in conversation_text:
                                for folder in accepted_folders:
                                    resource_folder = db_util.get_resource_path(folder['folder'].replace('folder-r', ''))
                                    folder['folder'] = resource_folder
                                
                                for folder in rejected_folders:
                                    resource_folder = db_util.get_resource_path(folder['folder'].replace('folder-r', ''))
                                    folder['folder'] = resource_folder
                                
                                template = get_final_confirmation_email_tempalte(owner_name, accepted_folders, rejected_folders)
                                response = util.send_reply(message_id, template)

                            elif "As confirmation, you are accepting Data Ownership for the following Managed Folders" in conversation_text and True:
                                for folder in accepted_folders:
                                    if status['accepted_verified'] and folder['email'] != 'NA' and folder['folder'].replace('folder-r', '') != '':
                                        resource_folder = db_util.get_resource_path(folder['folder'].replace('folder-r', ''))
                                        db_util.update_resource_accept(folder['email'], resource_folder)
                                
                                for folder in rejected_folders:
                                    if status['rejected_verified'] and folder['folder'].replace('folder-r', '') != '' :
                                        resource_folder = db_util.get_resource_path(folder['folder'].replace('folder-r', ''))
                                        name = folder['names'][0][0]
                                        deparment = folder['names'][0][1]
                                        title = folder['names'][0][2]
                                        suggested_owner = db_util.get_suggested_owner(name, deparment, title)
                                        result_status = db_util.update_resource_reject(email_id, resource_folder, status['suggested_owner_message_line'], suggested_owner, (name, deparment, title))
                except Exception as ex:
                    print(ex)
                    import sys
                    exc_type, exc_obj, exc_tb = sys.exc_info()
                    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                    print(exc_type, fname, exc_tb.tb_lineno)
                    continue               
    except Exception as ex:
        print(ex)
        import sys
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)





        
