import requests
import json
from urllib.parse import urlparse, parse_qs
from getpass import getpass
import configparser, os
config = configparser.ConfigParser()

config.read(os.path.join(os.path.join(os.path.join(os.path.dirname(os.path.dirname(__file__)),'api'),'config'),'config.ini'))
tenant = config["SHAREPOINT"]["tenant"]
site_url = f"https://{tenant}.sharepoint.com"

app_display_name = "Python API"
app_description = "Green Landscaping Python API"

# Replace with the redirect URI of your app (must be HTTPS)
redirect_uri = "https://localhost"

# Replace with the permissions that your app requires
app_permissions = {
    "list": ["Read", "Write"],
    "web": ["FullControl"]
}

registration_endpoint = site_url + "/_api/SP.AppRegNew"

app_registration_data = {
    "displayName": app_display_name,
    "description": app_description,
    "redirectUri": redirect_uri,
    "appPermissionRequests": []
}

# Add the app permissions to the app registration data
for resource, permissions in app_permissions.items():
    for permission in permissions:
        app_permission_request = {
            "Scope": resource,
            "Right": permission
        }
        app_registration_data["appPermissionRequests"].append(app_permission_request)

# Convert the app registration data to JSON
app_registration_json = json.dumps(app_registration_data)


def get_sharepoint_authorization_header_from_username_password(username,password):
    # Prompt the user to sign in to their Office 365 account
    if not username or not password:
        username = input("Enter your Office 365 email address: ")
        password = getpass("Enter your password: ")
    redirect_uri = "https://localhost"
    # Define the OAuth 2.0 authorization endpoint and parameters
    authorization_endpoint = "https://login.microsoftonline.com/common/oauth2/authorize"
    authorization_parameters = {
        "response_type": "code",
        "client_id": "",
        "redirect_uri": redirect_uri,
        "resource": site_url,
        "prompt": "login",
        "tenant": config["SHAREPOINT"]["tenant"]
    }
    print(json.dumps(authorization_parameters, indent=4))

    # Make a GET request to the authorization endpoint to obtain an authorization code
    authorization_url = authorization_endpoint + "?" + "&".join([f"{k}={v}" for k, v in authorization_parameters.items()])
    response = requests.get(authorization_url, auth=(username, password), allow_redirects=False)

    # Extract the authorization code from the response redirect URI
    print(response.headers)
    redirect_uri = response.headers["Location"]
    redirect_params = parse_qs(urlparse(redirect_uri).query)
    print(redirect_params)
    authorization_code = redirect_params["code"][0]

    # Define the OAuth 2.0 token endpoint and parameters
    token_endpoint = "https://login.microsoftonline.com/common/oauth2/token"
    token_parameters = {
        "grant_type": "authorization_code",
        "code": authorization_code,
        "redirect_uri": redirect_uri,
        "client_id": "", # config["SHAREPOINT"]["client_id"],
        "client_secret": "" # config["SHAREPOINT"]["client_secret"]
    }

    # Make a POST request to the token endpoint to obtain an access token
    response = requests.post(token_endpoint, data=token_parameters)

    # Parse the response JSON to get the access token
    response_json = json.loads(response.content)
    access_token = response_json["access_token"]

    # Define the headers for authenticated requests to the SharePoint REST API
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose"
    }
    return headers

# Make a POST request to register the app with SharePoint
# response = requests.post(registration_endpoint, headers=headers, data=app_registration_json)

# Print the response content
# print(response.content)
if __name__ == '__main__':
    with open(os.path.join(os.path.join(os.path.join(os.path.join(os.path.dirname(os.path.dirname(__file__)),'api'),'functions'),'Sharepoint'),'login.txt')) as f:
        p = eval(f.read())
        print(p["user"],p["pass"])
    print(get_sharepoint_authorization_header_from_username_password(p["user"],p["pass"]))