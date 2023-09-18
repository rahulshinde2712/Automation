__author__ = 'rahul.shinde'

"""
Mail Sending Script-with Attachment
"""

import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


class Mailtest:
    def __init__(self):
        self.employeeDtls = {}
        self.emailList = set()
        self.readingFiles()
        self.mail_send()



    def readingFiles(self):
        #Reading Files- List of items
        print('Reading employee Master file...')
        csvDictReader = csv.DictReader(open('AppData.txt'), delimiter='\t')
        self.employeeDtls = [dict(empid =row['EMPLOYEE_ID'], empemail=row['EMAIL'], name=row['EMP_NAME'] \
                                  ,mngname =row['MANAGER_EMAIL'] ,filename =row['FileName']) for row in csvDictReader]


    def mail_send(self):

        # emailList = self.uniqueEmailList()
        # employeeDtls = self.readingFiles()

        for r in self.employeeDtls:


            list1 = [str(r['empemail'])]

            print('Insurance Details Send to :' + str(r['empemail']))
            you = str(r['empemail']) + ";"

            cc = str(r['mngname'])
            copy = [str(r['mngname'])]
            # Create message container - the correct MIME type is multipart/alternative.
            msg = MIMEMultipart()
            msg['Subject'] = "TESTING:Letter_" + str(r['name'])
            msg['From'] = "xxx@test.com"
            msg['To'] = you
            msg['Cc'] = cc
            e_name = str(r['empemail']).split('.',1)
            html = '<html><head></head><body lang=EN-US link=blue vlink=purple>\
                           Hi ' + str(e_name[0]).capitalize() + ',<br><br>\
                           Body content.</body>\
                           <p>.......................................................<br><br>\
                           Regards,<br><br><span>Human Resource | <font color="blue">Aligned Automation </font></b><br><i>We Align before we Automate-so you can Accelerate</i><br><a href="https://www.test.com"><span>www.test.com</span></a><br><br><font size="2">This communication may contain privileged and/or confidential information. It is intended solely for the use of the recipient(s) named above. If you are not the intended recipient, you are strictly prohibited from disclosing, copying, distributing or using any of this information. If you received this communication in error, please contact the sender immediately and destroy the material in its entirety, whether electronic or hard copy.</font></p></div></html>'

            msgText = MIMEText(html, 'html')
            msg.attach(msgText)

            filename = str(r['filename']) + '.pdf'
            attachment = open('./Letters/' + str(r['filename']) + '.pdf', "rb")

            part = MIMEBase('application', 'octet-stream')
            part.set_payload((attachment).read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

            msg.attach(part)
            list_emp = list1 + copy
            s = smtplib.SMTP('smtp.office365.com', 587)
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login('mail@test.com', '****Password')
            s.sendmail('aa.hr@test.com', list_emp , msg.as_string())
            s.quit()


if __name__ == "__main__":

    #Calling class - Mail sending code
    run = Mailtest()