# MSgraph_api_python_sample
I had a very difficult time finding a working sample of code that would run API queries against AzureAD. This was used in a Jupyter Notebook, but should work anywhere.


To properly use this code, you need to replace the placeholder values for clientID, clientSecret, and tenantID with your actual application credentials. Make sure you have the necessary permissions and the appropriate endpoint to retrieve SharePoint site information.

The msgraph_auth() function authenticates with the Microsoft Graph API using client credentials, and the msgraph_request() function sends a GET request to the specified API endpoint.

Ensure that you have the required dependencies installed (msal, jwt, requests, and pandas) before running the code.
