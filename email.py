import imaplib
import os
import time 
import email

from email.utils import parseaddr
from email.header import decode_header
from email.parser import Parser

#import xlrd 
#import xlwt


mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login("xxxxx","xxxxxx")
mail.select(mailbox="INBOX", readonly=True)



type,data = mail.search(None, "(ALL)")


def catchData(msg):
    return msg["Date"]


def catchFrom(msg):
    mailFrom = msg["From"].split(" ")
    if len(mailFrom)==2:
        fromName = email.Header.decode_header((mailFrom[0]).strip('\"'))
        if fromName[0][1]:
            return unicode(fromName[0][0],fromName[0][1])+mailFrom[1]
        else:
            return unicode(fromName[0][0])+mailFrom[1]
    else:
        return msg["From"]

def catchSubject(msg):
    return msg["Subject"]

   

def mkdir(path):
 
    path=path.strip()
    path=path.rstrip("\\")
 
    isExists=os.path.exists(path)
 
    if not isExists:
       
        print(path+" Done")
    
        os.makedirs(path)
        return True
    else:
        
        print(path+" Path Existed")
        return False
 

    

ls = data[0].split()
ls.reverse()
for num in ls:
    #typ, data = mail.fetch(num, '(RFC822.HEADER)')
    #typ, data = mail.fetch(num, '(RFC822.TEXT)') #read all the text
    typ, data = mail.fetch(num, "(BODY.PEEK[HEADER])") 
    typ, data2 = mail.fetch(num,"(RFC822.TEXT)")
    
    mailText = data[0][1].decode("utf-8")
    mailText2 = data2[0][1].decode("utf-8")
    msg = email.message_from_string(mailText)
    textmsg = email.message_from_string(mailText2)
    mailDate = catchData(msg)
    mailFrom = catchFrom(msg)
    mailSubject = catchSubject(msg)
   
    

    if mailFrom==u"xxxxx xxxx <xxxxxx@xxxxx.com>":
    #if mailSubject[-8:]==u"The tips to get the most out of Gmail":    
    #if mailFrom==u"xxxxx xxxx <xxxxxx@xxxxx.com>" and mailSubject[-8:]==u"xxxxxxxx": # From%title must be exacly the same, email filter 
       
        
        
        f = open("record3.txt","a")
        print("Email Num:%s" % num,file=f)
        print("From:%s" % mailFrom,file=f)
        print("Date:%s" % mailDate,file=f)
        print("Subject:%s" % mailSubject,file=f)
        print(textmsg,file=f)
        f.close()
        
        

        path = "s:\\e"
        week = time.strftime('%W') 
        mkpath=path+u'\\ '+week+u"Week"
        mkdir(mkpath)
        
        print(u"Current path：%s" % path)

        try:
            if os.path.isdir(path):
                pass
            else:
                os.makedirs(mkpath)        
               
        except:
                print("Errorrrrr")
                f.close()

        book=xlwt.Workbook(encoding="utf-8",style_compression=0)       
        sheet = book.add_sheet(u"Sign-in",cell_overwrite_ok=False)
       
        #    sheet.write(l,c,mailSubject) not finished
           

        book.save(mkpath+u"\\Summary.xls")    
        


    else:
        print("not matched!!!")
        



mail.close()
mail.logout()











