# Python code to send email to a list of
# emails from a spreadsheet
  
# import the required libraries
import pandas as pd
import smtplib
  
# change these as per use
your_email = "kedaarvyas@gmail.com"
your_password = "ilwhqjgcndkmiute"
  
# establishing connection with gmail
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)
  
# reading the spreadsheet
email_list = pd.read_excel('C:/Users/kedaa/Downloads/emailfile.xlsx')
  
# getting the names and the emails
your_name="Kedaar Vyas"
names = email_list['NAME']
emails = email_list['EMAIL']
subjects=email_list['SUBJECT']
departments=email_list["DEPARTMENT"]
 
# iterate through the records
for i in range(len(emails)):
  
    # for every record get the name and the email addresses
    name = names[i]
    email = emails[i]
    subject=subjects[i]
    department=departments[i]
    # the message to be emailed
    message = "Hello Professor " + name + ",\n\nI came across your profile on the UMD Department of Computer Science website. I am very interested in your research in " + 
    department +". Your research topics are cutting edge and can be applied to various industries. I would like to be part of your research team during the 2022-2023 academic year.\n\nI am majoring in Computer Science-Machine Learning at UMD-College Park and have previous experience with languages such as Python, Java, and Matlab. I have also completed an undergraduate Computer Science programming course at the University of California-Berkeley (CS61A), where I applied Python to learn about topics such as object-oriented programming, linked lists, and recursion. I believe that I can use these skills to assist you in conducting your advanced research projects.\n\nPlease contact me at kedaarvyas@gmail.com or 484-744-7632.\n\nThank you,\nKedaar Vyas\nUndergraduate Degree student of CS at UMD-College Park CMNS department" 
  
    # sending the email
    full_email = ("From: {0} <{1}>\n"
                  "To: {2} <{3}>\n"
                  "Subject: {4}\n\n"
                  "{5}"
                  .format(your_name, your_email, name, email, subject, message))
                  
    server.sendmail(your_email, [email], full_email)
  
# close the smtp server
server.close()