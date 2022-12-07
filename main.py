import os, time 
import traceback
import numpy as np
import pandas as pd
from tqdm import tqdm
from imbox import Imbox

######################################################################################
### Utilities
######################################################################################
def create_or_check_storage(parent_path:str,path_to_save_email_content:str,path_to_save_attachment_file:str):   
    if not os.path.isdir(parent_path):
        os.makedirs(parent_path, exist_ok=True)

    if not os.path.isdir(path_to_save_email_content):
        os.makedirs(path_to_save_email_content, exist_ok=True)
    
    if not os.path.isdir(path_to_save_attachment_file):
        os.makedirs(path_to_save_attachment_file, exist_ok=True)

def save_attachment_file_in_directory(attachment_content,attachment_file_name,emailId,path_to_save_email_content):
    try:
        storage_path = os.path.join(path_to_save_email_content,f"{emailId}_{attachment_file_name}")
        if not os.path.isfile(storage_path):
            with open(storage_path,"wb") as fp:
                fp.write(attachment_content)  
        return storage_path     
    except:
        print(traceback.print_exc())

def save_email_content_as_html(email_content,emailId,path_to_save_attachment_file):
    try:
        storage_path = os.path.join(path_to_save_attachment_file,f"{emailId}_email.html")
        if not os.path.isfile(storage_path):
            with open(storage_path,"wb") as fp:
                fp.write(email_content)        
        return storage_path
    except:
        print(traceback.print_exc())

def create_csv(nested_list:list,csv_file_location:str):
    data_sheet = pd.DataFrame(np.array(nested_list),
                              columns=["email_numeric_id","sent_from","sent_to","subject","date","body_html_location","attach_file_location"])
    data_sheet.to_csv(csv_file_location,index=False)


######################################################################################
### main function
######################################################################################
def run_save(host:str, userId:str, pwd:str, parent_path:str,path_to_save_email_content:str,path_to_save_attachment_file:str, csv_file_location:str):
    try:
        # first check the storage location
        create_or_check_storage(parent_path,path_to_save_email_content,path_to_save_attachment_file)

        # starting the imbox session where we login to our account using app password
        with Imbox(host,username=userId,password=pwd,ssl=True,ssl_context=None,starttls=False) as imbox_session:

            # Selecting All emails 
            all_emails = imbox_session.messages(folder="all",uid__range='1:*')
            
            # initializing the nested list for creating a logging csv file 
            data = []
            for email_numeric_id, raw_email_obj in tqdm(all_emails):
                
                email_numeric_id = email_numeric_id.decode()
                
                subject_data = raw_email_obj.subject 
                subject_data = subject_data if subject_data else "no subject" 

                temp = [email_numeric_id,
                        raw_email_obj.sent_from[0]["email"],
                        raw_email_obj.sent_to[0]["email"],
                        subject_data,
                        raw_email_obj.date]
                
                # get body data from raw email
                email_body_data = raw_email_obj.body["html"]
                if len(email_body_data)>0:
                    email_body_path = save_email_content_as_html(email_body_data[0].encode(),
                                                                 email_numeric_id,
                                                                 path_to_save_email_content)
                    temp.append(email_body_path)
                else:
                    temp.append("body not attainable")

                # get the attachment(s) content for storage purpose
                attach_data = raw_email_obj.attachments
                if attach_data:
                    temp_attach_list = []
                    for attachment in attach_data:
                        attach_file_path = save_attachment_file_in_directory(attachment["content"].read(),
                                                                             attachment["filename"],
                                                                             email_numeric_id,
                                                                             path_to_save_attachment_file)
                        temp_attach_list.append(temp + [attach_file_path])
                    data.extend(temp_attach_list)
                else:
                    temp.append("no attachments")
                    data.append(temp)
            time.sleep(0.5)
        create_csv(data,csv_file_location)         
    except:
        print(traceback.print_exc())



if __name__ == "__main__":
    ######################################################################################
    ### Configuration
    ######################################################################################
    HOST = "imap.gmail.com" # using google's gmail her which can be changed.
    USER_ID = "**configure**" # Gmail Id.
    PWD = "**configure**" # after two-step verification, the app password should be used not the original one.

    # path to save email content, attachment files, csv file.
    PARENT_PATH = "Storage" # we can write any name here for parent directory, which will be created
    PATH_TO_SAVE_EMAIL_CONTENT = os.path.join(PARENT_PATH,"Emails") # works on both windows and mac(also linux) systems
    PATH_TO_SAVE_ATTACHMENT_CONTENT = os.path.join(PARENT_PATH,"Attachments")
    PATH_OF_CSV_FILE = os.path.join(PARENT_PATH,"log.csv") 

    ######################################################################################
    ### To run the process
    ######################################################################################
    run_save(host=HOST,
             userId=USER_ID,
             pwd=PWD,
             parent_path=PARENT_PATH,
             path_to_save_email_content=PATH_TO_SAVE_EMAIL_CONTENT,
             path_to_save_attachment_file=PATH_TO_SAVE_ATTACHMENT_CONTENT,
             csv_file_location=PATH_OF_CSV_FILE)






