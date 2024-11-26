from abc import ABC, abstractmethod

class AppWin(ABC):
  def __init__(self, appName,api):
    self._appName=appName
    self._api=api
  @abstractmethod
  def openNewApp(self):
    pass

import win32com.client as client
import time

class Outlook(AppWin):
  def __init__(self,appName, api, folderName):
    super().__init__(appName,api)
    self._folderName = folderName
    self.outlook = None
  def openNewApp(self):
    self.outlook = client.Dispatch(self._appName)
    namespace = self.outlook.GetNamespace(self._api)
    namespace.SendAndReceive(False)
    if self._folderName=="Inbox":
      self._folderNum = 6
      folder = namespace.GetDefaultFolder(self._folderNum)
    else:
      folder = namespace.GetDefaultFolder(5)
    return folder
  def quitOutlook(self):
    if self.outlook:
      self.outlook.Quit()
      print('Quit OK!')
if __name__ == "__main__":  
  outlookapp = Outlook("Outlook.Application", "MAPI", "Inbox")
  inbox = outlookapp.openNewApp()
  messages = inbox.Items
  print(len(messages))
  outlookapp.quitOutlook()





"""
import win32com.client as client
from datetime import datetime
import time

try:
    while True:
        # Tạo một instance của Outlook
        outlook = client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        namespace.SendAndReceive(False)
        # Truy cập thư mục inbox
        inbox = namespace.GetDefaultFolder(6)
        messages = inbox.Items
        lenMes1 = len(messages)
        print(len(messages))
        with open('messageCount.txt', 'r', encoding='utf-8') as file:
            lenMes2 = int(file.read())
            print(f"{lenMes2}, {lenMes1}")
        if lenMes2< lenMes1:
            sub = lenMes1 - lenMes2
            print(f"sub: {sub}")
            with open ('messageCount.txt','w' ) as file:
                file.write(str(lenMes1))
            list_mail = []
            for message in messages:
                list_mail.append((message.Subject, message.SenderEmailAddress , message.ReceivedTime))
            list_mail.sort(key=lambda x: x[2],reverse=True)
            # In ra danh sách đã sắp xếp
            i=0
            for subject, sender, received_time in list_mail:
                if i>=sub:
                    break
                print(f"Chủ đề: {subject}, Người gửi: {sender}, Thời gian nhận: {received_time}")
                i+=1
        time.sleep(30)
        
except Exception as e:
    print(f"Đã xảy ra lỗi: {e}")


#tìm theo Subject
import win32com.client as client

try:
    # Tạo một instance của Outlook
    outlook = client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")    
    # Truy cập thư mục inbox
    inbox = namespace.GetDefaultFolder(6)  # 6 là ID cho Inbox
    messages = inbox.Items
    # Chuỗi cần tìm kiếm
    search_string = "CÔNG TY CỔ PHẦN CÔNG NGHỆ TIN HỌC ANH NGỌC gửi hóa đơn điện tử số 00011583 cho CÔNG TY TNHH CENNOS ASIA"
    # Lặp qua tất cả email trong inbox
    found = False
    iter = 0
    for message in messages:
        # Kiểm tra xem tiêu đề email có chứa chuỗi tìm kiếm không
        if message.Subject and search_string in message.Subject:
            found = True
            print(f"Found in email: {message.Subject}, From: {message.SenderName}, Received: {message.ReceivedTime}, STT: {iter}")
        iter+=1
    if not found:
        print("Không tìm thấy email nào chứa chuỗi tìm kiếm.")
except Exception as e:
    print(f"Đã xảy ra lỗi: {e}")

import win32com.client as client
from datetime import datetime


outlook = client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
namespace = outlook.GetNamespace("MAPI")
# Lấy danh sách tất cả các tài khoản
for account in namespace.Accounts:
    print(f"Tài khoản: {account.DisplayName}")

inbox = namespace.GetDefaultFolder(6)
print(inbox)
messages = inbox.Items
count = 0
for i in messages:
  if count > 20:
    break
  count +=1
  print(i.Subject)




from lxml import html
tree = html.fromstring(messages[10].HTMLBody)
links = tree.xpath("//p")
for link in links:
     print(f"the P: {link.text_content()}")


the P: For FULL report check: https://intranet.cennext.com/basic/web/index.php?r=gmails/index
the P: These unverified emails will be REMOVED by IT TEAM if no further request.
the P: For more detail, check: https://intranet.cennext.com/basic/web/index.php?r=gmails/index
the P: This is automatically email, please DO NOT reply to this. Use IT Support Ticket feature if you need technical help.


import win32com.client as client

#You need to make sure that outlook must logged before and is closed before you run the script.

outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0) 
message.Display() 

"""
