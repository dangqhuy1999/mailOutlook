import win32com.client as client


#You need to make sure that outlook is closed before you run the script.
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0) 
message.Display() 