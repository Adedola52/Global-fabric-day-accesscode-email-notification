import os
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import numpy as np
import logging
import time
from dotenv import load_dotenv
import random
import string

logging.basicConfig(filename='microsoftfabric.log',
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s')

def microsoftfabric_access_code():
    """
        Loads registered users' data from an Excel file
     - Assigns a unique access code to each user
     - Saves the updated table with codes as a CSV file
     - Sends a confirmation email with the code via Gmail SMTP 
     - Includes both plain text and HTML versions of the email
     - Logs successes and errors appropriately

    """
    host = 'smtp.gmail.com'
    port = 587
    sender_email= 'fabriccommunitynigeria@gmail.com'
    fabric_pass = os.getenv('fabric_password')
    
    # reads excel file and save as a csv after access codes has been assigned 
    registration_form = pd.read_excel('Global Fabric Day 2025 (1-63).xlsx')
    registration_form['access_code'] = [''.join(random.choices(string.ascii_lowercase + string.digits, k=6)) for _ in range(len(registration_form))]
    registration_form.to_csv('fabric_day_access.csv', index=False)
    

    try:
        with smtplib.SMTP(host, port) as server:
            server.starttls()
            server.login(sender_email, fabric_pass)
            for _, row in registration_form.iterrows():
                try:
                    msg = MIMEMultipart('alternative')
                    msg['From'] = sender_email
                    msg['To'] = row['Email_Address']
                    msg['Subject'] = 'Global Fabric Day 2025 - Confirmation'
                    text = f""" Dear {row['Full_Name']},
                            Thank you for registering for *Global Fabric Day 2025*. We're excited to have you join us for this event. 

                            Date: 31st of May 2025  
                            Location:  Credit Direct, 48/50 Isaac John Street, Ikeja, Lagos State  
                            Your Access Code: {row['access_code']}

                            Please keep this code safe. It will be required at the check-in desk to access the event. Make sure to arrive on
                            time so you do not miss out on our amazing speaker sessions, hands-on demos, and community interactions.

                            If you have any questions, feel free to reach out to us at fabriccommunitynigeria@gmail.com.

                            See you there!

                            Warm regards,  
                            The Global Fabric Day Team
                        """
                    html= f"""
                            <html>
                            <body>
                            <p>Dear <strong>{row['Full_Name']}</strong>,</p>

                            <p>Thank you for registering for <strong>Global Fabric Day 2025</strong>. We're excited to have you join us for this event.</p>

                            <p>
                            <strong>Date:</strong> 31st of May 2025<br>
                            <strong>Location:</strong> Credit Direct, 48/50 Isaac John Street, Ikeja, Lagos State<br>
                            <strong>Your Access Code:</strong> {row['access_code']}
                            </p>

                            <p>Please keep this code safe. It will be required at the check-in desk to access the event. 
                            Make sure to arrive on time so you do not miss out on our amazing speaker sessions, hands-on demos, and community interactions.</p>

                            <p>If you have any questions, feel free to reach out to us at <a href="mailto:fabriccommunitynigeria@gmail.com">fabriccommunitynigeria@gmail.com</a>.</p>

                            <p>See you there!</p>

                            <p>Warm regards,<br>
                            The Global Fabric Day Team</p>
                            </body>
                            </html>
                            """

                    
                    text_message = MIMEText(text, 'plain')
                    html_message = MIMEText(html, 'html')
                    msg.attach(text_message)
                    msg.attach(html_message)


                    server.sendmail(sender_email, row['Email_Address'], msg.as_string())
                    logging.info(f"Access code sent to {row['Full_Name']} successfully")
                    time.sleep(1)

                # raises an exception when sending to a receipient fails
                except Exception as e:
                    logging.info(f"Email failed to send email to {row['Full_Name']}")
        

    
    # raises an exception when a connection error occurs 
    except smtplib.SMTPException as e:
            logging.error(f'Erro connecting to server')

microsoftfabric_access_code()
