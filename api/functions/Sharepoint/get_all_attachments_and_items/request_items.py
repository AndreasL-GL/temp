import requests
import json
import configparser
config = configparser.ConfigParser()
import os
import io
from PIL import Image
import datetime
config.read(os.path.join(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))),'config'),'config.ini'))

def get_sharepoint_access_headers_through_client_id():
    client_id = config["SHAREPOINT"]["client_id"].strip()
    client_secret = config["SHAREPOINT"]["client_secret"].strip()
    tenant_id = config["SHAREPOINT"]["tenant_id"].strip()
    tenant = config["SHAREPOINT"]["tenant"]
    client_id = client_id + '@'+tenant_id
    
    data = {
    'grant_type':'client_credentials',
    'resource': "00000003-0000-0ff1-ce00-000000000000/" + tenant + ".sharepoint.com@" + tenant_id, 
    'client_id': client_id,
    'client_secret': client_secret
}
    url = "https://accounts.accesscontrol.windows.net/tenant_id/tokens/OAuth/2"
    headers = {
    'Content-Type':'application/x-www-form-urlencoded'
}

    url = f"https://accounts.accesscontrol.windows.net/{tenant_id}/tokens/OAuth/2"
    r = requests.post(url, data=data, headers=headers)
    json_data = json.loads(r.text)
    headers = {
    'Authorization': "Bearer " + json_data['access_token'],
    'Accept':'application/json;odata=verbose',
    'Content-Type': 'application/json;odata=verbose'
}
    return headers

def get_all_items_one_week_back(site,list_):
    js = get_items(site,list_)
    items_this_week = [item for item in js if datetime.datetime.strptime(item['Modified'],'%Y-%m-%dT%H:%M:%SZ')>datetime.datetime.now()-datetime.timedelta(days=7) and item["Attachments"]]
    return items_this_week

def get_items(site, list_):
    headers=get_sharepoint_access_headers_through_client_id()
    tenant = config['SHAREPOINT']["tenant"]
    #"MKB Egenkontroll Augustenborg Periodiska 2023"
    url = f"https://{tenant}.sharepoint.com/sites/{site}/_api/web/lists/getbytitle('{list_}')/items"
    l = requests.get(url, headers=headers)
    js= json.loads(l.text)
    return js['d']['results']
    

def download_image(site, DecodedUrl,headers=get_sharepoint_access_headers_through_client_id()):
    tenant = config['SHAREPOINT']["tenant"]
    file_url = f"https://{tenant}.sharepoint.com/sites/{site}/_api/Web/GetFileByServerRelativeUrl('/{DecodedUrl}')/$value"
    file = io.BytesIO()
    file=io.BytesIO(requests.get(file_url,headers=headers).content)
    file.seek(0)
    #img = Image.open(file)
    #img.save(os.path.join(os.path.join(os.path.dirname(__file__),'images'),DecodedUrl.split('/')[-1]))
    return file
    
def get_urls_of_all_attachments(site = "GLMalmAB-EgenkontrollerGPISSVeloa",list_name = "Kommun Lekplatser Egenkontroll", one_week_back=False):
    headers=get_sharepoint_access_headers_through_client_id()
    if one_week_back: js= get_all_items_one_week_back(site,list_name)
    else: js = get_items(site,list_name)
    if len(js) ==0:return []
    if "Attachments" not in js[0].keys(): return []
    # Check for all items this past week.
    items_this_week = [item for item in js if datetime.datetime.strptime(item['Modified'],'%Y-%m-%dT%H:%M:%SZ')>datetime.datetime.now()-datetime.timedelta(days=7) and item["Attachments"]]
    file_list = []
    for item in items_this_week:
        uri = item['AttachmentFiles']['__deferred']['uri']
        rq = requests.get(uri,headers=headers)
        js = json.loads(rq.content)
        
        for nextitem in js['d']['results']:
            DecodedUrl = nextitem['ServerRelativePath']['DecodedUrl']
            file_list.append(DecodedUrl)
    return file_list

def download_all_files_one_week_back(site="GLMalmAB-EgenkontrollerGPISSVeloa", list_name="GP Egenkontroll_periodiska"):
    
    file_list = get_urls_of_all_attachments(site,list_name, one_week_back=True)
    return_list = []
    if not len(file_list)>10:
        print(len(file_list))
    else:
        for url in file_list:
            file = download_image(site,url,get_sharepoint_access_headers_through_client_id())
            return_list.append(file)
    return return_list

def download_all_attachments_in_list(site="GLMalmAB-EgenkontrollerGPISSVeloa", list_name="GP Egenkontroll_periodiska"):
    tenant = config['SHAREPOINT']['tenant']
    file_list = get_urls_of_all_attachments(site,list_name, one_week_back=False)
    return_list = []
    if not len(file_list)>10:
        print(len(file_list))
    else:
        for url in file_list:
            file = download_image(site,url,get_sharepoint_access_headers_through_client_id())
            
            return_list.append(file)
    return return_list
def get_all_attachment_links_from_site(site):
    """Gets all the links to attachments of a site to be used in requests.

    Args:
        site (_type_): _description_

    Returns:
        _type_: _description_
    """
    tenant = config["SHAREPOINT"]["tenant"]
    url =f"https://{tenant}.sharepoint.com/sites/{site}/_api/web/lists/"
    print(url)
    endpoint_url = url
    js = requests.get(endpoint_url, headers=get_sharepoint_access_headers_through_client_id())
    js = json.loads(js.content)['d']['results']
    titles = [item['Title'] for item in js]
    list_of_list_of_urls = []
    for item in titles: 
        print(item)
        list_of_list_of_urls.append(get_urls_of_all_attachments(site, item))
        print(list_of_list_of_urls)
    list_of_list_of_urls = [f"https://{tenant}.sharepoint.com/sites/{site}/_api/Web/GetFileByServerRelativeUrl('/{item}')/$value" for sublist in list_of_list_of_urls for item in sublist]
    
    return list_of_list_of_urls

if __name__=='__main__':
    js = get_all_attachment_links_from_site("GLMalmAB-EgenkontrollerGPISSVeloa")
    
    
    None
            
            
        
        
    
    

    # # SharePoint API endpoint to retrieve site information
    # api_url = f"{site_url}/_api/site?$select=ServerRelativeUrl"

    # response = requests.get(api_url, headers=headers)

    # # Get the server-relative URL from the response
    # server_relative_url = response.json()["d"]["ServerRelativeUrl"]
    # print(server_relative_url)