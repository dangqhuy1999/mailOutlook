import win32com.client as client

#You need to make sure that outlook must logged before and is closed before you run the script.

outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0) 
message.Display() 

"""
import win32com.client as client
outlook = client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)
messages = inbox.Items
print(messages[0].HTMLBody)

from lxml import html
tree = html.fromstring(messages[10].HTMLBody)
links = tree.xpath("//p")
for link in links:
     print(f"the P: {link.text_content()}")
the P: For FULL report check: https://intranet.cennext.com/basic/web/index.php?r=gmails/index
the P: These unverified emails will be REMOVED by IT TEAM if no further request.
the P: For more detail, check: https://intranet.cennext.com/basic/web/index.php?r=gmails/index
the P: This is automatically email, please DO NOT reply to this. Use IT Support Ticket feature if you need technical help.
"""
