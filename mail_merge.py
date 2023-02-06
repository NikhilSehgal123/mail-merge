import streamlit as st
import pandas as pd
import numpy as np
from ms_graph_api import *
import time

st.title('Vastmindz Mail Merge App')

access_token = None

if st.button('Sign in to Microsoft 365'):
    # Sign in to the Microsoft Graph API
    access_token = generate_access_token(APPLICATION_ID, SCOPES)
    st.success("Successfully signed in to the Microsoft Graph API!")

# Allow user to upload a CSV file
uploaded_file = st.file_uploader("Choose a CSV file", type="csv")

# Read the CSV file
if uploaded_file is not None:
    df = pd.read_csv(uploaded_file)
    
    # Drop any rows with missing values
    df = df.dropna()

    st.warning('We found ' + str(len(df)) + ' rows in the CSV file')

    st.header('Email Template')
    
    st.markdown("""
    Use the following placeholders in your email template:
    - {name} - The name of the person
    - {company} - The name of the company
    """)

    subject_template = "Engaging wellness solutions for {company}"

    subject_template = st.text_input('Subject', subject_template)

    outbound_email_template = st.text_area('Email Body')

    # Placeholder for slider
    slider_placeholder = st.slider('Time stagger for emails (seconds)', 0, 30, 20)

    if st.button('Send Emails'):
        st.write('Sending emails...')

        # Go though each row in the dataframe and send the email
        for index, row in df.iterrows():
            name = row['First name']
            email = row['Email']

            # Create the email body
            email_body = outbound_email_template.format(name=name, company=row['Company Name'])
            subject = subject_template.format(company=row['Company Name'])

            # Send the email
            send_email(subject, email_body, email)

            # Display a message
            st.write('Email sent to ' + email)

            # Wait for a few seconds
            time.sleep(slider_placeholder)