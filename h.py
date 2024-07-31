import win32com.client as win32 

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

y = input("Enter the number of spam busty wusty: ")
z = input("Enter the outlook email: ")

for x in range(1, y + 1):
    mailItem = olApp.CreateItem(0)
    mailItem.Subject = "Hello 123"
    mailItem.BodyFormat = 1
    mailItem.Body = "Hello There Noob"
    mailItem.To = z

    mailItem.Display()
    mailItem.Save()
    mailItem.Send()

    

# mailItem.BodyFormat = 2
# mailItem.HTMLBody = "<HTML Markup>"