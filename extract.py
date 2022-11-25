from Google import Create_Service
from googleapiclient.http import MediaIoBaseUpload
import io
import os
import sys
import base64
import time
import tkinter as tk
import tkinter.messagebox as msg
import re


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
    
def RemoveConnection():
    if os.path.exists(os.path.join(cwd, token_dir, gmailPickle)):
        os.remove(os.path.join(cwd, token_dir, gmailPickle))
    if os.path.exists(os.path.join(cwd, token_dir, drivePickle)):
        os.remove(os.path.join(cwd, token_dir, drivePickle))
    if os.path.exists(os.path.join(cwd, token_dir)):
        os.rmdir(os.path.join(cwd, token_dir))

    msg.showinfo("Connection Removed", "Existing email connection removed!")


def FolderNameSubmit():
        EntryFolderName = entry_foldername.get().strip()
        if EntryFolderName == "":
            msg.showinfo("Invalid", "Name cannot be empty.")
        elif len(EntryFolderName) > 40:
            msg.showinfo("Limit exceeded", "Maximum limit is 40 characters.")
        elif not EntryFolderName.isalnum():
            msg.showinfo("Invalid", "The name can only contain alphanumeric characters.")
        else:
            custom_parent_id = CreateDriveFolder(driveService, EntryFolderName)["id"]
            StoreAttachments([custom_parent_id])


def CustomFolderName():
    global entry_foldername
    label_folderquery.destroy()
    button_folderquery_yes.destroy()
    button_folderquery_no.destroy()

    label_foldername = tk.Label(root, text = "Enter the name of the folder:", font = ("Bahnschrift", 24), bg = "#ff9900")
    label_foldername.place(x = 400, y = 200)

    entry_foldername = tk.Entry(root, width = 70, font = ("Bahnschrift", 16))
    entry_foldername.place(x = 200, y = 300)

    button_foldername_submit = tk.Button(root, text = "Submit", command = FolderNameSubmit, padx = 10, pady = 10, bg = "#26ff00", activebackground = "#1fa308", font = ("Bahnschrift", 24), relief ='groove')
    button_foldername_submit.place(x = 532, y = 370)


def StoreAttachments(custom_parent_id = []):
    print("The attachments are being stored on your drive.")
    msg.showinfo("Extracting", "The attachments are being stored on your drive.")
    root.destroy()
    for emailMessage in emailMessages:
            messageId = emailMessage["threadId"]
            messageSubject = f"(No Subject) ({messageId})"
            messageDetail = GetMessageDetail(gmailService, emailMessage['id'], format='full', metadata_headers=['parts'])
            messageDetailPayload = messageDetail.get('payload')

            for item in messageDetailPayload["headers"]:
                if item["name"] == "Subject":
                    if item["value"]:
                        messageSubject = f"{item['value']} ({messageId})"
                        break

            folder_id = CreateDriveFolder(driveService, messageSubject, custom_parent_id)["id"]

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


def ExtractAttachments():
    global label_folderquery, button_folderquery_yes, button_folderquery_no, custom_parent_id, driveService, gmailService, emailMessages

    gmailService = ConstructService("gmail")
    if not os.path.exists(os.path.join(cwd, token_dir, drivePickle)):
        time.sleep(4)
    driveService = ConstructService("drive")

    button_extract.destroy()
    button_remove.destroy()

    querystring = "has:attachment"
    emailMessages = SearchEmail(gmailService, querystring, ["INBOX"])

    if emailMessages is None:
        print("No emails with attachments found.\nExiting...\n\n\n")
        msg.showinfo("Not found", "No emails with attachments found.")
        root.destroy()
        sys.exit()

    else:
        label_folderquery = tk.Label(root, text = "Would you like to create a custom folder for the attachments?", font = ("Bahnschrift", 24), bg = "#ff9900")
        label_folderquery.place(x=150, y = 200)
        button_folderquery_yes = tk.Button(root, text = "Yes", command = CustomFolderName, padx = 10, pady = 10, bg = "#26ff00", activebackground = "#1fa308", font = ("Bahnschrift", 24), relief ='groove')
        button_folderquery_yes.place(x = 450, y = 300)
        button_folderquery_no = tk.Button(root, text = "No", command = StoreAttachments, padx = 10, pady = 10, bg = "#ff0000", activebackground = "#8f0404", font = ("Bahnschrift", 24), relief ='groove')
        button_folderquery_no.place(x = 660, y = 300)


def Exit():
    print("Exiting...")
    root.destroy()
    sys.exit()


def ConstructTkinter():
    global root, button_extract, button_remove
    root = tk.Tk()
    X_dim, Y_dim = "1200", "600"
    root.geometry(f"{X_dim}x{Y_dim}")
    root.resizable(False, False)
    root.backGroundImage = tk.PhotoImage(file = "img/canvas.png")
    root.backGroundImageLabel = tk.Label(root, image = root.backGroundImage)
    root.backGroundImageLabel.place(x=0, y=0)
    root.iconImage = tk.PhotoImage(file = "img/drive.png")
    root.iconphoto(False, root.iconImage)
    root.title("Email Attachment Extractor")

    button_extract = tk.Button(root, text = "Extract Attachments", command = ExtractAttachments, padx = 10, pady = 10, bg = "#ff9900", activebackground = "#ad6903", font = ("Bahnschrift", 24), relief ='groove')
    button_extract.place(x=433, y = 100)
    button_remove = tk.Button(root, text = "Remove Existing Email", command = RemoveConnection, padx = 10, pady = 10, bg = "#ff9900", activebackground = "#ad6903", font = ("Bahnschrift", 24), relief ='groove')
    button_remove.place(x=418, y = 300)

    button_exit = tk.Button(root, text = "Exit", command = Exit, padx = 10, pady = 10, bg = "#ff3300", activebackground = "#ad3300", font = ("Bahnschrift", 24), relief ='groove')
    button_exit.place(x=555, y = 500)


def main():
    global cwd, token_dir, gmailPickle, drivePickle
    cwd = os.getcwd()
    token_dir = "token files"
    gmailPickle = "token_gmail_v1.pickle"
    drivePickle = "token_drive_v3.pickle"
    ConstructTkinter()
    root.mainloop()
    

if __name__ == "__main__":
    main()