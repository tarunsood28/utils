import smtplib
import os
import sys
import pandas as pd 


import mimetypes
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


COMMASPACE = ', '

class Email (object):
    
    def __init__ (self,server_details,username,password):
        self.server_details = server_details
        self.username=username
        self.password=password

    
    def _connect_to_smtp (self):
        _gmail_user = self.username
        _gmail_password = self.password
        try:  
            print ('creating connection')
            server = smtplib.SMTP(self.server_details)
            server.ehlo()
            server.starttls()
            server.login(_gmail_user, _gmail_password)
            return server
        except ConnectionResetError as e:  
            sys.exit('{}'.format(e))
    
    def create_msg (self,recepients,subject,email_body,attachments):
        outer  = MIMEMultipart()
        outer ['From']=self.username
        outer ['To']=COMMASPACE.join(recepients)
        outer ['Subject']=subject
        
        if email_body:
            for type,value in email_body.items():
                outer.attach(MIMEText(value,type))
        if attachments:
            for directory,files in attachments.items():            
                for filename in os.listdir(directory):
                    if not files or filename in files:
                        path = os.path.join(directory, filename)
                        if not os.path.isfile(path):
                            continue
                        # Guess the content type based on the file's extension.  Encoding
                        # will be ignored, although we should check for simple things like
                        # gzip'd or compressed files.
                        ctype, encoding = mimetypes.guess_type(path)
                        if ctype is None or encoding is not None:
                            # No guess could be made, or the file is encoded (compressed), so
                            # use a generic bag-of-bits type.
                            ctype = 'application/octet-stream'
                        maintype, subtype  = ctype.split('/', 1)
                        if maintype == 'text':
                            with open(path) as fp:
                                # Note: we should handle calculating the charset
                                msg = MIMEText(fp.read(), _subtype=subtype)
                        elif maintype == 'image':
                            with open(path, 'rb') as fp:
                                msg = MIMEImage(fp.read(), _subtype=subtype)
                        elif maintype == 'audio':
                            with open(path, 'rb') as fp:
                                msg = MIMEAudio(fp.read(), _subtype=subtype)
                        else:
                            with open(path, 'rb') as fp:
                                msg = MIMEBase(maintype, subtype)
                                msg.set_payload(fp.read())
                            # Encode the payload using Base64
                            encoders.encode_base64(msg)
                        msg.add_header('Content-Disposition', 'attachment', filename=filename)
                        outer.attach(msg)
        return outer
        
    def send_email (self,recepients,subject,email_body=None,attachments=None):
        server=self._connect_to_smtp () 
        try:
            server.send_message(self.create_msg(recepients,subject,email_body,attachments))
        except:
            print ('Something wrong happened here')
        finally:
            server.close()

if __name__ == '__main__':
    text = "Hi!\nHow are you?\nHere is the link you wanted:\nhttp://www.python.org"
    html = """\
    <html>
      <head></head>
      <body>
        <p>Regards!<br>
           Tarun Sood<br>
           <a href="http://www.python.org">Python Org</a>.
        </p>
      </body>
    </html>
    """
    email_body={}
    email_body['plain']=text
    email_body['html']=html
    mailer=Email('smtp.gmail.com:587','python.email9@gmail.com','Hellopython')
    mailer.send_email(['tarunsood69@gmail.com'],'Hello',email_body,{'E:\python':['sample.csv']})
    