
###Disclaimer
----------
### Following is a sample only for trying in local or development environment where users are not impacted. This is for reference only

### Step 1) The code was generated using github copilot as a boiler plate code
To reproduct, open VS code. Launch the  Github copilot chat window . Type the following prompt
create a databricks notebook code which has code to connect to sharepoint using certificate. it should be able to write to a list
SharePoint site: https://anishtest.sharepoint.com/sites/somelisttracker

### Step 2) above generated following code
## Install required libraries (run this cell first if not installed)
## pip install office365-rest-python-client azure-identity requests

### Step 3) create a databricks notebook and paste the code, review, make necessary changes for local/ dev env
### Databricks notebook source
##### Make neccsary changes and only try in local
from azure.identity import CertificateCredential   
from office365.sharepoint.client_context import ClientContext  
import requests  

#### COMMAND ----------

#### Configuration - Replace with your values
TENANT_ID = "your-tenant-id"   
CLIENT_ID = "your-client-id"   
CERTIFICATE_PATH = "/dbfs/path/to/your/certificate.pem"  # Path to certificate in DBFS  
SHAREPOINT_SITE_URL = "https://anishtest.sharepoint.com/sites/somelisttracker"  
LIST_NAME = "YourListName"  # Name of the SharePoint list  

#### COMMAND ----------  

#### Authenticate using certificate  
credential = CertificateCredential(  
    tenant_id=TENANT_ID,  
    client_id=CLIENT_ID,  
    certificate_path=CERTIFICATE_PATH  
)  

#### Get access token
token = credential.get_token("https://anishtest.sharepoint.com/.default")  
access_token = token.token  

#### COMMAND ----------

#### Connect to SharePoint site
ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(credential)  

#### COMMAND ----------

#### Get the list
list_obj = ctx.web.lists.get_by_title(LIST_NAME)  

#### COMMAND ----------

#### Example: Add an item to the list  
#### Replace with your actual field names and values  
item_properties = {  
    "Title": "Sample Item",  
    "Field1": "Value1",  
    "Field2": "Value2"  
}  

list_item = list_obj.add_item(item_properties)   
ctx.execute_query()  

print(f"Item added with ID: {list_item.id}")  

#### COMMAND ----------

#### Optional: Query existing items
items = list_obj.items.get().execute_query()  
for item in items:  
    print(item.properties)  

