import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from constants import password

# Path to your resume
resume_path = "Kumar_Utsav_Resume.pdf"  # Replace with the actual file path
recruiter_list_path = "recruiters_list.xlsx"  # Replace with the actual file path


def load_recruiter():
    recruiter_df = pd.read_excel(recruiter_list_path, sheet_name="Sheet1")
    return recruiter_df


# Email details
def set_email_details(rec_name, rec_email, comp_name, comp_detail):
    from_email = "kutsav056.btech2021@ee.nitrr.ac.in"
    to_email = rec_email
    subject = f"Application for Full-Time Role in Software Development at {comp_name}"
    body = f"""
    Dear {rec_name},

    I hope this message finds you well.

    My name is Kumar Utsav, and I am currently pursuing  B.Tech in Electrical Engineering at the National Institute of Technology Raipur, with an expected graduation date in 2025. I am reaching out to express my interest in potential full-time roles within {comp_name} that align with my skills and academic background.

    Throughout my academic career, I have maintained a CGPA of 9.39 out of 10. I have developed a strong foundation in core CS subjects such OOPS, Computer Networks, Operating System and DBMS. I have full-stack development, with extensive knowledge of Flutter, Express, and Django, which I have utilized to develop full-stack 
    applications. Additionally, I possess introductory knowledge in machine learning and deep learning, and have successfully developed a plant disease classifier using these technologies.

    My experience also includes contributing as a mobile app developer to various committees at my college, where I have honed my skills in software development and team collaboration.

    I am particularly drawn to {comp_name} because of {comp_detail}. I am confident that my background and skills would make me a valuable addition to your team.

    I have attached my resume for your review and would welcome the opportunity to discuss how my experience and skills can contribute to the continued success of {comp_name}. Additionally, you can view my portfolio at https://uts-folio.netlify.app/.

    Thank you for considering my application. I look forward to the possibility of speaking with you soon.

    Warm regards,

    Kumar Utsav,
    Final Year Student,
    National Institute of Technology Raipur,
    Mob No.: 7250241229
    Email id: kumarutsav.9434@gmail.com
    """

    return from_email, to_email, subject, body


# SMTP server configuration
def server_setup(from_email):
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure the connection
        server.login(from_email, password)
        print("Login Successful")
        return server
    except smtplib.SMTPAuthenticationError as e:
        print("Failed to send email due to authentication error:", e)


# Create a multipart message and set headers
def send_email(from_email, to_email, subject, body, server):
    message = MIMEMultipart()
    message["From"] = from_email
    message["To"] = to_email
    message["Subject"] = subject

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    # Attach resume
    attachment = MIMEBase("application", "octet-stream")
    try:
        with open(resume_path, "rb") as file:
            attachment.set_payload(file.read())
        encoders.encode_base64(attachment)
        attachment.add_header(
            "Content-Disposition",
            f"attachment; filename={os.path.basename(resume_path)}"
        )
        message.attach(attachment)
    except Exception as e:
        print(f"Failed to attach resume {resume_path}: {e}")

    try:
        # Connect to the server
        server.sendmail(from_email, to_email, message.as_string())
        print("Email sent successfully!")
    except Exception as e:
        print("Failed to send email:", e)


def main_handler():
    server = server_setup('kutsav056.btech2021@ee.nitrr.ac.in')
    recruiter_df = load_recruiter()
    for index, row in recruiter_df.iterrows():
        if row['Is Sent'] == "NO":
            from_email, to_email, subject, body = set_email_details(row['Recruiter Name'], row['Recruiter Email'],
                                                                    row['Company Name'], row['About Company'])
            send_email(from_email, to_email, subject, body, server)

    server.quit()


if __name__ == "__main__":
    main_handler()
