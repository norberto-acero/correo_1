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

#%% Mostrar los correos no leídos

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


#%% #! los correos por conversaciones--------------------------------------------------
import win32com.client as client

# Dirección de correo del usuario (ajústala si es diferente)
tu_direccion = "norberto.acero@ejemplo.com".lower()

# Iniciar Outlook
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')

# Acceder a la bandeja de entrada y luego a la subcarpeta "NACERO"
inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
nacero_folder = inbox.Folders['NACERO']

# Obtener los correos y ordenarlos por fecha
items = nacero_folder.Items
items.Sort("[ReceivedTime]", True)

# Diccionario para agrupar los correos por conversación
conversaciones = {}

# Recorrer los correos
for i in range(1, items.Count + 1):
    try:
        mail = items.Item(i)
        if mail.Class != 43:  # Solo MailItems
            continue

        # Verificar que no estás en el campo 'To'
        if tu_direccion not in str(mail.To).lower():
            topic = mail.ConversationTopic
            if topic not in conversaciones:
                conversaciones[topic] = []
            conversaciones[topic].append(mail)

    except Exception:
        continue  # Ignora errores (p.ej., si hay ítems no de correo)

# Mostrar conversaciones con más de un mensaje
for topic, mails in conversaciones.items():
    if len(mails) > 1:
        print(f"\n🧵 Conversación: {topic}")
        for mail in mails:
            print(f"  Asunto: {mail.Subject}")
            print(f"  De: {mail.SenderName}")
            print(f"  Para: {mail.To}")
            print(f"  CC: {mail.CC}")
            print(f"  Fecha: {mail.ReceivedTime}")
            print("-" * 50)

#%% #? todas las conversaciones con correos pendintes--------------------------------------------------
import win32com.client as client

# Tu dirección de correo (ajústala si es diferente)
tu_direccion = "norberto.acero@ejemplo.com".lower()

# Iniciar Outlook
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')

# Acceder a la subcarpeta 'NACERO' dentro de la bandeja de entrada
inbox = namespace.GetDefaultFolder(6)
nacero = inbox.Folders['NACERO']

# Obtener solo los correos no leídos
unread_items = nacero.Items.Restrict("[Unread] = true")
unread_items.Sort("[ReceivedTime]", True)

# Agrupar por conversación, filtrando los que NO están dirigidos a ti
conversaciones = {}

for i in range(1, unread_items.Count + 1):
    try:
        mail = unread_items.Item(i)
        if mail.Class != 43:  # Solo MailItem
            continue

        # Si tu correo NO está en el campo 'To'
        if tu_direccion not in str(mail.To).lower():
            topic = mail.ConversationTopic
            if topic not in conversaciones:
                conversaciones[topic] = []
            conversaciones[topic].append(mail)

    except Exception:
        continue  # Ignora errores (ej. ítems corruptos)

# Mostrar conversaciones con más de un mensaje
for topic, mails in conversaciones.items():
    if len(mails) > 1:
        print(f"\n🧵 Conversación: {topic}")
        for mail in mails:
            print(f"  Asunto: {mail.Subject}")
            print(f"  De: {mail.SenderName}")
            print(f"  Para: {mail.To}")
            print(f"  CC: {mail.CC}")
            print(f"  Fecha: {mail.ReceivedTime}")
            print("-" * 50)
