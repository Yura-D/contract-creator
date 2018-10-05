from __future__ import print_function
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from httplib2 import Http
from oauth2client import client, tools, file
import io

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/drive'

def main():
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('client_id.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('drive', 'v3', http=creds.authorize(Http()))

    # Call the Drive v3 API
    results = service.files().list(
        pageSize=10, fields="nextPageToken, files(id, name)", q="'1fGkO8AuLMJIN9gL-oVxWfZNS2rTQfMnp' in parents").execute() # you can change text 'test' "name contains 'test'"
    items = results.get('files', [])

    if not items:
        print('No files found.')
    else:
        print('Files:')
        for item in items:
            print(item['id'])
            print('{0} ({1})'.format(item['name'], item['id']))

    """
    # upload file
        folder_id = '1fGkO8AuLMJIN9gL-oVxWfZNS2rTQfMnp'
        file_metadata = {'name': 'photo.jpeg',
                        'parents': folder_id}
        media = MediaFileUpload('photo.jpeg',
                                mimetype='image/jpeg')
        file_upload = service.files().create(body=file_metadata,
                                            media_body=media,
                                            fields='id').execute()
        print('File ID: %s' % file_upload.get('id'))
        
    """
    """
    #download file  
    
    file_id = '1G7AwGl4TidnX0ksy3QYL1Ay6cfzwWsjN'
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))

    with io.open('1ph.jpeg', 'wb') as f:
        fh.seek(0)
        f.write(fh.read())
    """

if __name__ == '__main__':
    main()