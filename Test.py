# import streamlit as st
# import requests
# import urllib.parse

# # Azure AD App details
# client_id = 'dcc0356c-a07e-4272-9b1c-cb635e0fc2a0'
# client_secret = 'G1x8Q~XtNmetzdEE_2cOyyFl84YXlLzLjEV~caQZ'
# tenant_id = 'db01513b-9352-48dc-9232-8fc9f4e6979f'
# redirect_uri = 'http://localhost:8501/'

# # Endpoints
# authority = f'https://login.microsoftonline.com/{tenant_id}'
# # authorize_url = f'{authority}/oauth2/v2.0/authorize'
# token_url = f'{authority}/oauth2/v2.0/token'

# # Authorization parameters
# scope = 'openid email profile'
# response_type = 'id_token'

# def get_authorization_url():
#     params = {
#         'client_id': client_id,
#         'response_type': response_type,
#         'redirect_uri': redirect_uri,
#         'response_mode': 'form_post',
#         'scope': scope
#     }
#     fou = f'{token_url}?{urllib.parse.urlencode(params)}'
#     print(fou)
#     return fou

# def get_token_from_code(code):
#     headers = {
#         'Content-Type': 'application/x-www-form-urlencoded'
#     }
#     body = {
#         'grant_type': 'authorization_code',
#         'code': code,
#         'redirect_uri': redirect_uri,
#         'client_id': client_id,
#         'client_secret': client_secret,
#         'scope': scope
#     }
#     response = requests.post(token_url, headers=headers, data=body)
#     print(response.json())
#     return response.json()

# def fetch_user_data(token):
#     headers = {
#         'Authorization': f'Bearer {token}'
#     }
#     response = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers)
#     return response.json()

# if 'token' not in st.session_state:
#     auth_url = get_authorization_url()
#     st.write(f'Click [here]({auth_url}) to log in.')

#     code = st.experimental_get_query_params().get('code')
#     if code:
#         code = code[0]
#         print(code)
#         token_response = get_token_from_code(code)
#         if 'access_token' in token_response:
#             st.session_state.token = token_response['access_token']
#         else:
#             st.error('Authentication failed.')

# if 'token' in st.session_state:
#     token = st.session_state.token
#     user_data = fetch_user_data(token)
#     st.write("Logged in as:")
#     st.json(user_data)


import streamlit as st
import requests
import urllib.parse

# Azure AD App details
client_id = 'dcc0356c-a07e-4272-9b1c-cb635e0fc2a0'
client_secret = 'G1x8Q~XtNmetzdEE_2cOyyFl84YXlLzLjEV~caQZ'
tenant_id = 'db01513b-9352-48dc-9232-8fc9f4e6979f'
redirect_uri = 'http://localhost:8501/'

# Endpoints
authority = f'https://login.microsoftonline.com/{tenant_id}'
token_url = f'{authority}/oauth2/v2.0/token'

# Authorization parameters
scope = 'openid email profile'

# def get_token_from_code(code):
#     headers = {
#         'Content-Type': 'application/x-www-form-urlencoded'
#     }
#     body = {
#         'grant_type': 'authorization_code',
#         'code': code,
#         'redirect_uri': redirect_uri,
#         'client_id': client_id,
#         'client_secret': client_secret,
#         'scope': scope
#     }
#     response = requests.post(token_url, headers=headers, data=body)
#     return response.json()


# if 'token' not in st.session_state:
#     code = st.experimental_get_query_params().get('code')
#     if code:
#         code = code[0]
#         token_response = get_token_from_code(code)
#         if 'access_token' in token_response:
#             st.session_state.token = token_response['access_token']
#         else:
#             st.error('Authentication failed. Please ensure you provided the correct credentials and try again.')
#     else:
st.write(f'Please authenticate by visiting the following URL: {authority}/oauth2/v2.0/token?client_id={client_id}&response_type=id_token&redirect_uri={urllib.parse.quote(redirect_uri)}&response_mode=form_post&scope={urllib.parse.quote(scope)}')

