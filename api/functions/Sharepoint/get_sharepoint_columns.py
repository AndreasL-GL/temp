import json, os
from shareplum import Site
from shareplum import Office365
import requests

with open(os.path.join(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))),'config'),'sharepoint_egenkontroller_remove_list.txt')) as f:
    sharepoint_columns_remove_list = [x.replace('\n','') for x in f.readlines() if x != '\n']


def get_body_from_sharepoint_api(js):
    resultlist = []
    if 'body' in js.keys(): js = js['body']
    for item in js['d']['results']:
        if item['StaticName'] not in sharepoint_columns_remove_list:
            resultlist.append({"Moment":item['Title'], "link":item['StaticName']})
    return resultlist
def get_sharepoint_access_headers():
    client_id = "50d23ac1-8de9-4941-9399-004472826045"
    client_secret = "1Zo5LkdK4ZZfzuIXKO2o2FslnZBC5keyqqpEqW1YWEs="
    tenant_id = "a096cfba-db7b-4c9c-9506-d8e91da824ee"
    tenant = "greenlandscapingmalmo"
    client_id = client_id + '@'+tenant_id
    site = "GLMalmAB-EgenkontrollerVellingebostder"
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
def get_fields(site, list_):
    headers=get_sharepoint_access_headers()
    tenant = "greenlandscapingmalmo"
    #site = "GLMalmAB-EgenkontrollerVellingebostder"
    url = f"https://{tenant}.sharepoint.com/sites/{site}/_api/web/lists/getbytitle('{list_}')/fields"
    print(url)
    l = requests.get(url, headers=headers)
    js= json.loads(l.text)
    js = get_body_from_sharepoint_api(js)
    return js
