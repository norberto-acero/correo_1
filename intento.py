#%%
import win32com.client
from datetime import datetime

# Conectar con Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Obtener la bandeja de entrada y el subfolder "NACERO"
inbox = namespace.GetDefaultFolder(6)  # 6 es la carpeta de la Bandeja de Entrada
nacero_folder = inbox.Folders["NACERO"]

# Obtener los correos electrónicos en el subfolder "NACERO"
messages = nacero_folder.Items
messages.Sort("[ReceivedTime]", True)  # Ordenar por fecha de recepción, más recientes primero

# Filtrar los correos electrónicos que no están enviados directamente a ti y que están sin leer
unread_messages = []

for msg in messages:
    try:
        # Verificar si el correo no ha sido leído
        if msg.UnRead:
            # Comprobar si el correo no fue enviado directamente a ti
            if not msg.To or namespace.CreateRecipient(msg.To).Address != namespace.CurrentUser.Address:
                # Filtrar mensajes de la misma conversación (agrupados)
                conversation = msg.ConversationID
                if conversation not in [m.ConversationID for m in unread_messages]:
                    unread_messages.append(msg)
    except Exception as e:
        print(f"Error procesando el mensaje: {e}")

# Mostrar los mensajes que cumplen con los filtros
if unread_messages:
    print(f"Se han encontrado {len(unread_messages)} mensajes sin leer que no fueron enviados directamente a ti y están agrupados.")
    for message in unread_messages:
        print(f"De: {message.SenderName}, Asunto: {message.Subject}, Recibido: {message.ReceivedTime}")
else:
    print("No se encontraron mensajes que cumplan con los criterios.")
