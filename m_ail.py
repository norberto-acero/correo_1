#%% #? toda la carpeta--------------------------------------------------
import win32com.client as client

# Iniciar Outlook
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')

# Acceder a la Bandeja de entrada
inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

# Acceder a la subcarpeta llamada "NACERO"
subfolder = inbox.Folders['NACERO']

# Listar los correos en la subcarpeta
items = subfolder.Items
items.Sort("[ReceivedTime]", True)  # Ordenar por fecha de recepción (descendente)
items.Count


#%% Mostrar el último correo
if items.Count > 0:
    message = items.Item(1)
    print("Remitente:", message.SenderName)
    print("Asunto:", message.Subject)
    print("Cuerpo:", message.Body[:500])  # Solo los primeros 500 caracteres
else:
    print("La subcarpeta 'NACERO' está vacía.")


#%% #TODO correos no leidos--------------------------------------------------
import win32com.client as client

# Iniciar Outlook y acceder al espacio de nombres MAPI
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')

# Acceder a la Bandeja de entrada
inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

# Acceder a la subcarpeta 'NACERO'
subfolder = inbox.Folders['NACERO']

# Obtener solo los correos no leídos
unread_items = subfolder.Items.Restrict("[Unread] = true")
unread_items.Sort("[ReceivedTime]", True)  # Orden descendente por fecha
print(unread_items.Count)

# Mostrar los correos no leídos

if unread_items.Count > 0:
    print(f"Hay {unread_items.Count} correos no leídos en 'NACERO':\n")
    for i in range(1, unread_items.Count + 1):
        mail = unread_items.Item(i)
        print(f"Asunto: {mail.Subject}")
        print(f"De: {mail.SenderName}")
        print(f"Fecha: {mail.ReceivedTime}")
        print(f"Vista previa del cuerpo:\n{mail.Body[:200]}")
        print("-" * 50)
else:
    print("No hay correos no leídos en 'NACERO'.")
