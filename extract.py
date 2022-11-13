from Google import Create_Service
from googleapiclient.http import MediaIoBaseUpload
import io
import os
import sys
import base64
import time

def ConstructService(apiService):
    ClientServiceFile = "Client_Secret_Attachment_Extractor.json"
    try:
        if apiService == "drive":
            apiName = "drive"
            apiVersion = "v3"
            Scopes = ["https://www.googleapis.com/auth/drive"]
            return Create_Service(ClientServiceFile, apiName, apiVersion, Scopes)
        elif apiService == "gmail":
            apiName = "gmail"
            apiVersion = "v1"
            Scopes = ["https://mail.google.com/"]
            return Create_Service(ClientServiceFile, apiName, apiVersion, Scopes)
    except Exception as e:
        print(e)
        return None

def SearchEmail(service, querystring, label_ids=[]):
        messageResponse = service.users().messages().list(userId="me", labelIds=label_ids, q=querystring).execute()
        
        messageItems = messageResponse.get("messages")
        nextPageToken = messageResponse.get("nextPageToken")

        while nextPageToken:
            messageResponse = service.users().messages().list(userId="me", labelIds=label_ids, q=querystring, pageToken=nextPageToken).execute()

            messageItems.extend(messageResponse.get('messages'))
            nextPageToken = messageResponse.get('nextPageToken')
        return messageItems

def GetMessageDetail(service, message_id, format="metadata", metadata_headers=[]):
    try:
        messageDetail = service.users().messages().get(userId="me",id=message_id,format=format,metadataHeaders=metadata_headers).execute()
        return messageDetail

    except Exception as e:
        print(e)
        return None

def CreateDriveFolder(service, folderName, parentFolder=[]):
    file_metadata = {"name": folderName, "parents": parentFolder, "mimeType": "application/vnd.google-apps.folder"}
    file = service.files().create(body=file_metadata, fields="id").execute()
    return file
    
def main():

    cwd = os.getcwd()
    token_dir = "token files"
    gmailPickle = "token_gmail_v1.pickle"
    drivePickle = "token_drive_v3.pickle"
    while True:

        print("1. Connect to new email")
        print("2. Use connected email")
        print("3. Exit program")
        opt = input("Enter your option: ").strip()
        if opt == "1":
            if os.path.exists(os.path.join(cwd, token_dir, gmailPickle)):
                os.remove(os.path.join(cwd, token_dir, gmailPickle))
            if os.path.exists(os.path.join(cwd, token_dir, drivePickle)):
                os.remove(os.path.join(cwd, token_dir, drivePickle))
            if os.path.exists(os.path.join(cwd, token_dir)):
                os.rmdir(os.path.join(cwd, token_dir))
            break
        elif opt == "2":
            if not os.path.exists(os.path.join(cwd, token_dir)):
                print("No connected email found, creating new connection...")
            break
        elif opt == "3":
            print("Exiting...\n\n\n")
            sys.exit()
        else:
            print("Invalid input, try again.\n\n")

    gmailService = ConstructService("gmail")
    time.sleep(2)
    driveService = ConstructService("drive")

    querystring = "has:attachment"
    emailMessages = SearchEmail(gmailService, querystring, ["INBOX"])
    
    if emailMessages is None:
        print("No emails with attachments found.\n\n\n")
        main()
    else:
        custom_parent_id = None
        while True:
            opt_custom = input("Would you like to create a custom folder for the attachments? (Y/N)").lower()
            if opt_custom == "y":
                while True:
                    custom_folder_name = input("Enter the name of the folder: ")
                    if custom_folder_name.isalnum():
                        custom_parent_id = CreateDriveFolder(driveService, custom_folder_name)["id"]
                        break
                    else:
                        print("Folder name can only be alphanumeric, try again.\n")
                break
            elif opt_custom == "n":
                break
            else:
                print("Invalid input, try again\n")
        for emailMessage in emailMessages:
            messageId = emailMessage["threadId"]
            messageSubject = f"(No Subject) ({messageId})"
            messageDetail = GetMessageDetail(gmailService, emailMessage['id'], format='full', metadata_headers=['parts'])
            messageDetailPayload = messageDetail.get('payload')
            # print(messageDetailPayload)
            for item in messageDetailPayload["headers"]:
                if item["name"] == "Subject":
                    if item["value"]:
                        messageSubject = f"{item['value']} ({messageId})"
                    else:
                        messageSubject = f"(No Subject) ({messageId})"

            folder_id = CreateDriveFolder(driveService, messageSubject, [] if custom_parent_id is None else [custom_parent_id])["id"]

            if "parts" in messageDetailPayload:
                for msgPayload in messageDetailPayload["parts"]:
                    mimeType = msgPayload["mimeType"]
                    fileName = msgPayload["filename"]
                    body = msgPayload["body"]

                    if "attachmentId" in body:
                        attachment_id = body["attachmentId"]

                        response = gmailService.users().messages().attachments().get(userId="me", messageId=emailMessage["id"], id=attachment_id).execute()
                        fileData = base64.urlsafe_b64decode(response.get('data').encode('utf-8'))
                        fh = io.BytesIO(fileData)
                        fileMetadata = {"name": fileName, "parents": [folder_id]}
                        mediaBody = MediaIoBaseUpload(fh, mimetype=mimeType, chunksize=1024 * 1024, resumable=True)

                        file = driveService.files().create(body=fileMetadata, media_body=mediaBody, fields="id").execute()

if __name__ == "__main__":
    main()