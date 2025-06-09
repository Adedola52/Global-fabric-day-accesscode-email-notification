# Global Fabric Day 2025 – Access Code Email Notification

As a volunteer for the Microsoft Fabric Nigeria User Group, I had the opportunity to give back to the community by applying my skills to support the success of Global Fabric Day 2025.

## Problem Statement
Attendees who registered for the event were to be sent a personalized access code in other to access the event. To solve this:

- I collected attendee data using Microsoft Forms
- I generated unique access codes (a mix of text and numbers) using Python
- The codes were saved to a .csv file for easy reference and tracking
  
## Tool Used:
- Python

## Solution Overview
I created a function that:

- Reads the Excel file exported from Microsoft Forms
- Generates unique access codes
- Saves the updated data (including codes) as a CSV
- Sends personalized emails to each attendee using the data
  
The function leverages smtplib to connect to the email server and uses MIMEMultipart to send both plain text and HTML versions of the email body, ensuring compatibility with different email clients

The function also includes error handling to ensure that any issues during the email-sending process or server connection are logged and managed properly.

## Email Recipients Received:
![image alt](https://github.com/Adedola52/Global-fabric-day-accesscode-email-notification/blob/e20cac8a21d683d33f1a12f788ccc7b881bf8c22/Fabric%20Day%20Access%20code%20email.jpg)

