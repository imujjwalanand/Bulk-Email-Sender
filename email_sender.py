import smtplib
from email.mime.multipart import MIMEMultipart   # this is a specific import for Python3, for python2, the import goes like email.MIMEMultipart
from email.mime.text import MIMEText    # the same import instruction as above
from email.mime.base import MIMEBase	# the same import instruction as above
from email import encoders	
from openpyxl import load_workbook

workbook = load_workbook('emails.xlsx')   # open the file that contains all the details of the company like email address, name of person, name of company etc. 

first_sheet = workbook.get_sheet_names()[0]
worksheet = workbook.get_sheet_by_name(first_sheet) # this is the main sheet under consideration

fromaddr = "imujjwalanand@gmail.com"
email_range = worksheet['A1':'A4']
company_range = worksheet['B1':'B4']   #Change the starting and ending cell values for the variables according to yourself. 


# Now here, there are two variables email and company name, you can add or remove the variables according to your choice. 



for cell1, cell2 in zip(email_range, company_range):    # looping through all the email ids and recording the variables
	for email, company in zip(cell1, cell2):
		toaddr = email.value
		company_name = company.value
		 
		msg = MIMEMultipart()
		 
		msg['From'] = fromaddr      # from EMAIL ADDRESS
		msg['To'] = toaddr			# to EMAIL ADDRESS (This one is picked from the xlsx file)
		msg['Subject'] = "Looking for full time employment opportunity"   # the SUBJECT line
		 
		body = "Hello " + company_name + ". Greetings from Ujjwal Anand here! \n Hope you have a good day!\n Regards\nUjjwal Anand\nIIT Jodhpur"  # the message body

		#Can use the \n to draft your message body and make it look real 

		msg.attach(MIMEText(body, 'plain'))
		 
		filename = "Resume_Ujjwal_Anand.pdf"  # write the name of the file that you want to attach
		attachment = open("Resume_Ujjwal_Anand.pdf", "rb") # here write the pathname of the file that you want to attach
		 
		part = MIMEBase('application', 'octet-stream')
		part.set_payload((attachment).read())
		encoders.encode_base64(part)
		part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
		 
		msg.attach(part)
		 
		server = smtplib.SMTP('smtp.gmail.com', 587)
		server.starttls()
		server.login(fromaddr, "YOUR PASSWORD")     # DON'T FORGET TO LOGIN, ENTER THE PASSWORD HERE
		text = msg.as_string()
		server.sendmail(fromaddr, toaddr, text)
server.quit()     # always quit/close the server once you're done. 
		