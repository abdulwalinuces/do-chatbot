
import os
import openai
import re
import json


# openai.api_type = "azure"
# openai.api_base = "https://ebdev.openai.azure.com/"
# openai.api_version = "2023-05-15"
# openai.api_key = "***************"


# openai.api_type = 'azure'
# openai.api_version = '2023-05-15' # this may change in the future

# deployment_name= 'StandardModel' #'gpt35_16k_Model' 

openai.api_key = '' 

def get_prompt(conversation ):
    return  f"""Given the email conversation below, identify the owners of each folder. Provide a JSON list structure with fields: 'folder', 'name', and 'email' for each folder. Enclose JSON property name and value in double quotes. Get the actual folder path as folder. Do not use the folder number. If a piece of information is not available from the conversation, use "NA" as a placeholder. If a suggested name is a team, group, or similar entity, extract only the person's name from it.
                Email Conversation:
                {conversation}
                """
    
def process_prompt(prompt):
    try:
        # print(prompt, '\n', '='*100)
        message=[{"role": "user", "content": prompt}]
        response = openai.ChatCompletion.create(
                                                    model="gpt-3.5-turbo",
                                                    messages = message,
                                                    temperature=0.2,
                                                    max_tokens=1500,
                                                    frequency_penalty=0.0
                                                )
        #text = response['choices'][0].text.replace(' .', '.').strip()
        text = response['choices'][0]['message']['content']
        return str(text)
        
    
    except Exception as ex:
        return ''
    
def get_accepted_resources_verifiactio_prompt(email_conversation):
    """
    Reconfirms the accepted resources by the owner. Which the OpenAI model has found in the email conversation.
    """
    prompt =  f"""Email Conversation:
            {email_conversation}
            Please analyze the emails in the conversation and determine if the owner has verified the lists of accepted and rejected resources. Also get the one line from email in which the suggested owner name is mentioned by the owner. Provide the response in JSON format:"""
    prompt +=  """{
                "accepted_verified": "True/False",
                "rejected_verified": "True/False",
                "suggested_owner_message_line": "the line in which suggested owner name is mentioned."
                }"""
    return prompt
