"""
Demonstrates how to authenticate with user credentials (username and password) in non-interactive mode


"""
from office365.sharepoint.client_context import ClientContext
#from tests import test_password, test_site_url, test_username
from office365.runtime.auth.user_credential import UserCredential

username = 'robotFicheros'
password = 'Sencilla1'
base_url = 'https://totemtowersspain.sharepoint.com/'
site_url = '/sites/Prueba_QLIK'

test_username=username
test_password=password
test_site_url=base_url+site_url

ctx = ClientContext(test_site_url)
ctx.with_user_credentials(test_username, test_password)

web = ctx.web.get().execute_query()
print(web.url)