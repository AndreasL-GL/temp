from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
import os
import configparser
APP_BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
config = configparser.ConfigParser()
config.read(os.path.join(os.path.join(APP_BASE_DIR,'config'),'config.ini'))
# site_url: URL of your SharePoint site
# username: username to authenticate with
# password: password to authenticate with
# app_principal_id: the client ID of your Azure AD application
# app_principal_secret: the client secret of your Azure AD application
# xml_content: the XML content to grant app permissions
site_url = "https://greenlandscapingmalmo.sharepoint.com/sites/GLMalmAB-EgenkontrollerMalmstad"
with open(os.path.join(os.path.dirname(__file__), 'login.txt')) as f: p=eval(f.read())
authcookie = Office365(site_url, username=p["user"], password=p["pass"]).GetCookies()
xml_content="""<AppPermissionRequests AllowAppOnlyPolicy="true">
  <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="FullControl" />
</AppPermissionRequests>"""

site = Site(site_url, version=Version.v365, authcookie=authcookie)

app_principal_id = config["SHAREPOINT"]["client_id"]
# Give app permissions to the site
site.grant_app_principal_permission(app_principal_id, "Site Collection Administrator", xml_content=xml_content)

# Add the app principal as a site collection administrator
site.add_site_collection_administrator(app_principal_id)

# Clear app permissions for the site
site.clear_app_principal_permission(app_principal_id)