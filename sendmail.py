import win32com.client as win
ol=win.Dispatch("outlook.application")
olmailitem=0x0 #size of the new email
newmail=ol.CreateItem(olmailitem)
newmail.Subject= 'Testing Mail'
newmail.To='sathishx.ramalingareddy@intel.com'
newmail.CC='sathishx.ramalingareddy@intel.com'
#newmail.CC='naveenasivani@gmail.com'
#newmail.CC='shanmukhx.sb@intel.com'
newmail.Body= 'Please ignore, this is a test email to oneboxdeployment@outlook.com showcase how to send auto-emails using python.'
# attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
# newmail.Attachments.Add(attach)
# To display the mail before sending it
# newmail.Display()
newmail.Send()