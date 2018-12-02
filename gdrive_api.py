from __future__ import print_function
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from httplib2 import Http
from oauth2client import client, tools, file
import io
import configuration

# If modifying these scopes, delete the file token.json.
# Authorization to Google drive with personal token
SCOPES = 'https://www.googleapis.com/auth/drive'
store = file.Storage(configuration.drive_token)
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets(configuration.client_id, SCOPES)
    creds = tools.run_flow(flow, store)
service = build('drive', 'v3', http=creds.authorize(Http()))


def get_list(gfolder):
    # Call the Drive v3 API
    # get list of some folder that you need

    results = service.files().list(
        fields="nextPageToken, files(id, name)", q= "'{0}' in parents".format(gfolder)).execute() 
                                # you can change text 'test' "name contains 'test'"
    items = results.get('files', [])

    
    if not items:
        print('No files found.')
    else:
        folder_list = {}
        item_number = 0
        for item in items:
            item_number += 1
            folder_list[item_number] = [item['name'], item['id']]
    return folder_list

def gupload(gfolder, gfile, mimetype, path=""):
# upload file
       
    file_metadata = {'name': [gfile],
                    'parents': [gfolder]}
    media = MediaFileUpload(path+gfile,
                            mimetype=mimetype)
    file_upload = service.files().create(body=file_metadata,
                                        media_body=media,
                                        fields='id').execute()
    print('File ID: %s' % file_upload.get('id'))


def gdrive_search(gsearch):
    # Call the Drive v3 API
    results = service.files().list(
        fields="nextPageToken, files(id, name)", q= "name contains '{0}'".format(gsearch)).execute()
    items = results.get('files', [])

    if not items:
        print('No files found.')
    else:
        print('Files:')
        for item in items:
            print('* {0} - ({1})'.format(item['name'], item['id']))


def gdownload(file_id, named_file):
    # download file binary file
    
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))

    with io.open(named_file, 'wb') as f:
        fh.seek(0)
        f.write(fh.read())


def doc_download(file_id, named_file, write_path):

    # download docx file
    request = service.files().export_media(fileId=file_id,
                                            mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    # fh = io.BytesIO()
    fh = io.FileIO(write_path + named_file, mode='wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))
    """
    with io.open(named_file, 'wb') as f:
        fh.seek(0)
        f.write(fh.read())"""

    'temp\\'