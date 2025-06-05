# Global Fabric Day 2025 – Access Code Email Notification

As a volunteer for the Microsoft Fabric Nigeria User Group, I had the opportunity to give back to the community by applying my skills to support the success of Global Fabric Day 2025.

## Problem Statement
Attendees who registered for the event via Microsoft Forms needed to receive a personalized access code to join the event. To solve this:

- I collected attendee data using Microsoft Forms
- I generated unique access codes (a mix of text and numbers) using Python
- The codes were saved to a .csv file for easy reference and tracking
  
## Tool Used:
- Python

## Solution Overview
I created a modular Python function that:

- Reads the Excel file exported from Microsoft Forms
- Generates unique access codes
- Saves the updated data (including codes) as a CSV
- Sends personalized emails to each attendee using the data
  
The function leverages smtplib to connect to the email server and uses MIMEMultipart to send both plain text and HTML versions of the email body — ensuring compatibility with different email clients

The function also includes error handling, so any exception during the email-sending process is logged or handled gracefully.
