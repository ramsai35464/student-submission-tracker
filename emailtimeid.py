import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from datetime import datetime, timedelta
from apscheduler.schedulers.background import BackgroundScheduler

# Function to send emails
def send_reminder_emails(not_submitted, sender_email, sender_password, deadline):
    try:
        # Setup SMTP server connection
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)

        # Loop through students and send email
        for _, row in not_submitted.iterrows():
            recipient_email = row['Email']
            student_id = row['ID']  # Use ID instead of Name

            subject = "Submission Deadline Missed"
            body = f"""
            Dear {student_id},  # Use student ID for the email

            You have missed the submission deadline of {deadline}.
            Please complete your work and submit it as soon as possible.

            Best regards,
            Admin
            """

            msg = MIMEText(body)
            msg['Subject'] = subject
            msg['From'] = sender_email
            msg['To'] = recipient_email

            server.sendmail(sender_email, recipient_email, msg.as_string())
            st.write(f"Reminder email sent to {student_id} at {recipient_email}")

        # Close SMTP server connection
        server.quit()
        st.success("Reminder emails sent successfully!")
    except Exception as e:
        st.error(f"An error occurred while sending emails: {e}")

# Streamlit application title
st.title("Student Submission Tracker with Deadline")

# Section to upload Excel files
st.header("Upload Excel Files")
all_students_file = st.file_uploader("Upload All Students Excel File", type=["xlsx"])
submitted_students_file = st.file_uploader("Upload Submitted Students Excel File", type=["xlsx"])

# Separate inputs for deadline date and time
deadline_date = st.date_input("Set the Submission Deadline (Date)")
deadline_time = st.time_input("Set the Submission Deadline (Time)")

# Combine date and time into a single datetime object
if deadline_date and deadline_time:
    deadline = datetime.combine(deadline_date, deadline_time)
else:
    deadline = None

# Process only when both files are uploaded and a deadline is set
if all_students_file and submitted_students_file and deadline:
    try:
        # Load data from uploaded files into pandas DataFrames
        all_students = pd.read_excel(all_students_file)
        submitted_students = pd.read_excel(submitted_students_file)

        # Validate column names
        if 'ID' not in all_students.columns or 'Email' not in all_students.columns:
            st.error("The 'All Students' Excel file must contain 'ID' and 'Email' columns.")
        elif 'ID' not in submitted_students.columns:
            st.error("The 'Submitted Students' Excel file must contain an 'ID' column.")
        else:
            # Identify students who have not submitted
            not_submitted = all_students[~all_students['ID'].isin(submitted_students['ID'])]

            # Display students who have not submitted
            st.subheader("Students Who Have Not Submitted")
            st.dataframe(not_submitted)

            if not not_submitted.empty:
                st.write(f"Total students who have not submitted: {len(not_submitted)}")

                # Input fields for email credentials
                sender_email = st.text_input("Enter your email (Gmail address)")
                sender_password = st.text_input("Enter your App Password", type="password")  # App password for Gmail

                # Schedule emails at the deadline
                if st.button("Schedule Emails"):
                    scheduler = BackgroundScheduler()
                    send_time = deadline

                    # Schedule the send_reminder_emails function
                    scheduler.add_job(
                        send_reminder_emails,
                        'date',
                        run_date=send_time,
                        args=[not_submitted, sender_email, sender_password, deadline]
                    )
                    scheduler.start()
                    st.success(f"Emails scheduled to be sent at {send_time}")
    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
else:
    st.info("Please upload both Excel files and set the deadline to continue.")
