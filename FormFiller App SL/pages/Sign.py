import os
from streamlit_extras.switch_page_button import switch_page
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader
import streamlit as st

# Determine the directory of the script
script_directory = os.path.dirname(os.path.abspath(__file__))
config_file_path = os.path.join(script_directory, 'config.yaml')

# Load configuration from the config file
with open(config_file_path) as file:
    config = yaml.load(file, Loader=SafeLoader)

# Initialize the authenticator with configuration
authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days'],
    config['preauthorized']
)

# Perform authentication
name, authentication_status, username = authenticator.login(fields={'Login': 'main'})

# Handle authentication status
if st.session_state["authentication_status"]:
    authenticator.logout('Logout', 'main', key='unique_key')
    st.title(f'Welcome *{st.session_state["name"]}*')
    switch_page("app")
elif st.session_state["authentication_status"] is False:
    st.error('Username/password is incorrect')
elif st.session_state["authentication_status"] is None:
    st.warning('Please enter your username and password')

# Save the updated configuration back to the config file
with open(config_file_path, 'w') as file:
    yaml.dump(config, file, default_flow_style=False)
