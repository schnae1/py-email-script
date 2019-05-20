# Reads emails from Excel spredsheet
# Sends predefined emails to list in file
# Utilizes OpenPyXL for reading from excel
# Uses yagmail API to send emails

import openpyxl
import yagmail

# Open excel 
wb = openpyxl.load_workbook('emails.xlsx')

# Get excel sheet
sheet1 = wb.active

for row in sheet1.values:
	for value in row:
		receiver = value
		body = "Hi to all, \n This is a test format."
		#filename = "document.pdf"

		# from email
		yag = yagmail.SMTP("email_address@gmail.com")
		yag.send(
    			to=receiver,
    			subject="Yagmail test with attachment",
    			contents=body, 
    			#attachments=filename,
		)
		print(value)
