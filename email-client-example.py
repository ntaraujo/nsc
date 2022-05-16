import sys
import os

import smtplib
import mimetypes

from email import encoders
from email.utils import formatdate
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from PySide6 import QtWidgets, QtCore, QtGui

QtWidgets.QDialog

attachments = []
labels = []

class Login(QtWidgets.QDialog):
    def __init__(self,parent = None):
        QtWidgets.QDialog.__init__(self, parent)
        
        self.initUI()

    def initUI(self):

        self.live = QtWidgets.QRadioButton("Windows Live",self)
        self.gmail = QtWidgets.QRadioButton("Google Mail",self)
        self.yahoo = QtWidgets.QRadioButton("Yahoo! Mail",self)

        e = [self.live,self.gmail,self.yahoo]

        for i in e:
            i.clicked.connect(self.Email)

        self.userl = QtWidgets.QLabel("E-mail:",self)
        self.user = QtWidgets.QLineEdit(self)

        self.passl = QtWidgets.QLabel("Password:",self)
        
        self.passw = QtWidgets.QLineEdit(self)
        self.passw.setEchoMode(self.passw.Password)

        self.echo = QtWidgets.QCheckBox("Show/Hide password",self)
        self.echo.stateChanged.connect(self.Echo)

        self.go = QtWidgets.QPushButton("Login",self)
        self.go.clicked.connect(self.Login)

        grid = QtWidgets.QGridLayout()

        grid.addWidget(self.live,0,0)
        grid.addWidget(self.gmail,0,1)
        grid.addWidget(self.yahoo,0,2)
        grid.addWidget(self.userl,1,0,1,1)
        grid.addWidget(self.user,1,1,1,2)
        grid.addWidget(self.passw,2,1,1,2)
        grid.addWidget(self.passl,2,0,1,1)
        grid.addWidget(self.echo,3,0,1,2)
        grid.addWidget(self.go,3,2)

        self.setLayout(grid)

        self.setGeometry(300,300,350,200)
        self.setWindowTitle("PyMail Login")
        self.setWindowIcon(QtGui.QIcon("PyMail"))
        self.setStyleSheet("font-size:15px;")

    def Echo(self,state):
        if state == QtCore.Qt.Checked:
            self.passw.setEchoMode(self.passw.Normal)
        else:
            self.passw.setEchoMode(self.passw.Password)

    def Email(self):
        global account
        account = self.sender().text()

    def Login(self):
        global account
        global server
        global user

        user = self.user.text()

        if account == "Windows Live":
            server = smtplib.SMTP('smtp.live.com',25)

        elif account == "Google Mail":
            server = smtplib.SMTP('smtp.gmail.com',25)

        elif account == "Yahoo! Mail":
            server = smtplib.SMTP('smtp.mail.yahoo.com',465)

        try:    
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(user, self.passw.text())

            self.hide()

            main = Main(self)
            main.show()
            
        except smtplib.SMTPException:
            msg = QtWidgets.QMessageBox.critical(self, 'Login Failed',
            "Username/Password combination incorrect", QtWidgets.QMessageBox.Ok | 
            QtWidgets.QMessageBox.Retry, QtWidgets.QMessageBox.Ok)

            if msg == QtWidgets.QMessageBox.Retry:
                self.Login()
        
class Main(QtWidgets.QMainWindow):

    def __init__(self,parent=None):
        QtWidgets.QMainWindow.__init__(self,parent)
        self.initUI()

    def initUI(self):
        global user

        self.send = QtWidgets.QPushButton("Send",self)
        self.send.clicked.connect(self.Send)

        self.from_label = QtWidgets.QLabel("From",self)

        self.to_label = QtWidgets.QLabel("To",self)

        self.subject_label = QtWidgets.QLabel("Subject",self)

        self.from_addr = QtWidgets.QLineEdit(self)
        self.from_addr.setText(user)

        self.to_addr = QtWidgets.QLineEdit(self)
        self.to_addr.setPlaceholderText("godfather@corleone.it")

        self.subject = QtWidgets.QLineEdit(self)
        self.subject.setPlaceholderText("I got an offer you can't refuse")

        self.image = QtWidgets.QPushButton("Attach file",self)
        self.image.clicked.connect(self.Image)
        
        self.text = QtWidgets.QTextEdit(self)

        centralwidget = QtWidgets.QWidget()

        self.grid = QtWidgets.QGridLayout()

        self.grid.addWidget(self.from_label,0,0)
        self.grid.addWidget(self.from_addr,1,0)
        self.grid.addWidget(self.to_label,2,0)
        self.grid.addWidget(self.to_addr,3,0)
        self.grid.addWidget(self.subject_label,4,0)
        self.grid.addWidget(self.subject,5,0)
        self.grid.addWidget(self.image,6,0)
        self.grid.addWidget(self.text,8,0)
        self.grid.addWidget(self.send,9,0)

        centralwidget.setLayout(self.grid)

        self.setCentralWidget(centralwidget)


#---------Window settings --------------------------------
        
        self.setGeometry(300,300,500,500)
        self.setWindowTitle("PyMail")
        self.setWindowIcon(QtGui.QIcon("PyMail"))
        self.setStyleSheet("font-size:15px")

    def ContextMenu(self):
        global sender
        sender = self.sender()
        
        self.menu = QtWidgets.QMenu(self)

        remove = QtWidgets.QAction("Remove",self)
        remove.triggered.connect(self.Remove)

        self.menu.addAction(remove)

        self.menu.show()

    def Remove(self):
        global sender
        global pos
        global labels
        global attachments

        pos -= 1

        ind = labels.index(sender)

        attachments.remove(attachments[ind])

        labels.remove(sender)

        sender.setParent(None)

    def Image(self):
        global path
        global attachments
        global labels
        global filetype
        global l

        path = QtWidgets.QFileDialog.getOpenFileName(self, "Attach file","/home/")[0]

        if path:
            
            attachments.append(path)

            filetype = path[path.rindex(".")+1:]

            if filetype == "png":
                pic = QtGui.QPixmap(path)
            else:
                if filetype+".png" in os.listdir("C:/Python32/python/pyqt/PyMail/48px/"):
                    print("normal")
                    pic = QtGui.QPixmap("C:/Python32/python/pyqt/PyMail/48px/"+filetype+".png")
                else:
                    print("weird")
                    pic = QtGui.QPixmap("C:/Python32/python/pyqt/PyMail/48px/_blank.png")
                    
            a = QtWidgets.QLabel(path,self)
            a.setScaledContents(True)
            a.setFixedSize(50,50)
            a.setPixmap(pic)
            a.setToolTip(path)

            a.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
            a.customContextMenuRequested.connect(self.ContextMenu)

            labels.append(a)

            print(attachments,labels)
            
            pos = len(attachments)

            l = [self.from_label,self.from_addr,self.to_label,self.to_addr,self.subject_label,self.subject,self.image,self.text,self.send]
            
            for index,i in enumerate(l):
                self.grid.addWidget(i,index,0,1,pos+1)

                if i in l[-2:]:
                    self.grid.addWidget(i,index+1,0,1,pos+1)

            self.grid.addWidget(a,7,pos-1)
            self.setGeometry(300,300,500,550)
        
    def Send(self):
        global server
        global attachments
        global filetype
        global l
        
        fromaddr = self.from_addr.text()
        toaddr = self.to_addr.text()
        subject = self.subject.text()
        
        msg = MIMEMultipart()
        msg['From'] = fromaddr
        msg['To'] = toaddr
        msg['Subject'] = subject
        msg['Date'] = formatdate()

        body = self.text.toPlainText()
        msg.attach(MIMEText(body,"plain"))

        if attachments:
            for file in attachments:

                ctype, encoding = mimetypes.guess_type(file)

                if ctype is None or encoding is not None:
                    ctype = 'application/octet-stream'
                    
                maintype, subtype = ctype.split('/', 1)

                if maintype == 'text':
                    fp = open(file)
                    att = MIMEText(fp.read(), _subtype=subtype)
                    fp.close()
                elif maintype == 'image':
                    fp = open(file, 'rb')
                    att = MIMEImage(fp.read(), _subtype=subtype)
                    fp.close()
                elif maintype == 'audio':
                    fp = open(file, 'rb')
                    att = MIMEAudio(fp.read(), _subtype=subtype)
                    fp.close()
                else:
                    fp = open(file, 'rb')
                    att = MIMEBase(maintype, subtype)
                    att.set_payload(fp.read())
                    fp.close()
                    encoders.encode_base64(att)

                att.add_header('Content-Disposition', 'attachment', filename=file[file.rindex("/"):])
                msg.attach(att) 

        text = msg.as_string()
        
        try:
            server.sendmail(fromaddr, toaddr, text)

            msg = QtWidgets.QMessageBox.information(self, 'Message sent',
            "Message sent successfully, clear everything?", QtWidgets.QMessageBox.Yes | 
            QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.Yes)

            if msg == QtWidgets.QMessageBox.Yes:
                self.to_addr.clear()
                self.subject.clear()
                self.text.clear()

                if attachments:
                    for i in attachments:
                        attachments.remove(i)

                    for i in reversed(range(self.grid.count())):
                        self.grid.itemAt(i).widget().setParent(None)

                    for index,i in enumerate(l):
                        self.grid.addWidget(i,index,0)
            
        except smtplib.SMTPException:
            
            msg = QtWidgets.QMessageBox.critical(self, 'Error',
            "The message could not be sent, retry?", QtWidgets.QMessageBox.Yes | 
            QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.Yes)

            if msg == QtWidgets.QMessageBox.Yes:
                self.Send()
        
def main():
    app = QtWidgets.QApplication(sys.argv)
    login = Login()
    login.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()