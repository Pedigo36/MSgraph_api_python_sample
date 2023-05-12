import msal
import jwt
import json 
import requests 
import pandas as pd 
from datetime import datetime


accessToken = None 
requestHeaders = None 
tokenExpiry = None 
queryResults = None 
graphURI = 'https://graph.microsoft.com'

def msgraph_auth():
    """
    Authenticates with the Microsoft Graph API using client credentials.
    """
    global accessToken
    global requestHeaders
    global tokenExpiry
    
    #for more information onthese fields:
    #https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app
    clientID = "{cliendID}"
    clientSecret = "{clientSecret}"
    tenantID = "{tenantID}"
    authority = 'https://login.microsoftonline.com/' + tenantID
    scope = ['https://graph.microsoft.com/.default']
    
    app = msal.ConfidentialClientApplication(clientID, authority=authority, client_credential=clientSecret)

    try:
        accessToken = app.acquire_token_silent(scope, account=None)
        if not accessToken:
            try:
                accessToken = app.acquire_token_for_client(scopes=scope)
                if accessToken['access_token']:
                    print('New access token retrieved....')
                    requestHeaders = {'Authorization': 'Bearer ' + accessToken['access_token']}
                else:
                    print('Error acquiring authorization token. Check your tenantID, clientID, and clientSecret.')
            except:
                pass 
        else:
            print('Token retrieved from MSAL Cache....')

        decodedAccessToken = jwt.decode(accessToken['access_token'], verify=False)
        accessTokenFormatted = json.dumps(decodedAccessToken, indent=2)
        print('Decoded Access Token')
        print(accessTokenFormatted)

        # Token Expiry
        tokenExpiry = datetime.fromtimestamp(int(decodedAccessToken['exp']))
        print('Token Expires at: ' + str(tokenExpiry))
        return
    except Exception as err:
        print(err)

def msgraph_request(resource, requestHeaders):
    """
    Sends a GET request to the Microsoft Graph API.
    """
    # Request
    results = requests.get(resource, headers=requestHeaders).json()
    return results

# Auth - this executes the auth request
msgraph_auth()



# Query
# to run additional queries, update the URL with the different Graph endpoints
# this is a great resource to find the API you would like: https://developer.microsoft.com/en-us/graph/graph-explorer
# this specific query will list all sharepoints available in your environment. You must have proper permissions to run this query
queryResults = msgraph_request('https://graph.microsoft.com/v1.0/sites?search=*', requestHeaders)

# Results to Dataframe
try:
    df = pd.read_json(json.dumps(queryResults['value']))
    # Set ID column as index
    df = df.set_index('id')
    print(df['displayName'])  # Assuming the 'displayName' field contains the site name

except:
    print(json.dumps(queryResults, indent=2))
