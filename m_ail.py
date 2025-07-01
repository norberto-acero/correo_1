#%%
import win32com.client as client
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
inbox = namespace.GetDefaultFolder(6)
subfolder = inbox.Folders['NACERO']
items = subfolder.Item
if items.Count > 0:
    message = items.Item(1)
    print("Remitente:", message.SenderName)
    print("Asunto:", message.Subject)
    print("Cuerpo:", message.Body[:500])
else:
    print("La subcarpeta 'NACERO' está vacía.")
