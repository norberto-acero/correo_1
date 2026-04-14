import win32com.client as client

# ===== SETUP =====
def get_folder(folder_name="NACERO"):
    outlook = client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)
    return inbox.Folders[folder_name]

def get_sorted_items(items):
    items.Sort("[ReceivedTime]", True)
    return items

# ===== 1. UNREAD EMAILS BY CONVERSATION =====
def get_unread_conversations(folder, user_email):
    unread_items = get_sorted_items(folder.Items.Restrict("[Unread] = true"))

    conversations = {}

    for i in range(1, unread_items.Count + 1):
        try:
            mail = unread_items.Item(i)
            if mail.Class != 43:
                continue

            if user_email not in str(mail.To).lower():
                topic = mail.ConversationTopic
                conversations.setdefault(topic, []).append(mail)

        except Exception:
            continue

    return conversations

def print_unread_conversations(conversations):
    for topic, mails in conversations.items():
        mails.sort(key=lambda m: m.ReceivedTime, reverse=True)

        print(f"\n🧵 Conversation: {topic}")
        for mail in mails:
            print(f"Subject: {mail.Subject}")
            print(f"From: {mail.SenderName}")
            print(f"To: {mail.To}")
            print(f"Date: {mail.ReceivedTime}")
            print("-" * 40)

# ===== 2. LIST UNREAD EMAILS =====
def get_unread_emails(folder):
    return get_sorted_items(folder.Items.Restrict("[Unread] = true"))

def print_unread_emails(unread_items):
    count = unread_items.Count
    print(f"\nUnread emails: {count}\n")

    for i in range(1, count + 1):
        mail = unread_items.Item(i)
        print(f"Subject: {mail.Subject}")
        print(f"From: {mail.SenderName}")
        print(f"Date: {mail.ReceivedTime}")
        print(f"Preview: {mail.Body[:200]}")
        print("-" * 40)

# ===== 3. SEARCH EMAIL BY SUBJECT =====
def search_email_by_subject(folder, subject_keyword):
    items = get_sorted_items(folder.Items)

    results = []

    for i in range(1, items.Count + 1):
        try:
            mail = items.Item(i)
            if mail.Class == 43 and subject_keyword.lower() in mail.Subject.lower():
                results.append(mail)
        except Exception:
            continue

    results.sort(key=lambda m: m.ReceivedTime, reverse=True)
    return results


def print_search_results(results):
    print(f"\nResults found: {len(results)}\n")

    for mail in results:
        print(f"Subject: {mail.Subject}")
        print(f"From: {mail.SenderName}")
        print(f"Date: {mail.ReceivedTime}")
        print(f"Preview: {mail.Body[:200]}")
        print("-" * 40)

# ===== MAIN USAGE =====
if __name__ == "__main__":
    USER_EMAIL = "norberto.acero@ejemplo.com".lower()
    folder = get_folder("NACERO")

    unread_conversations = get_unread_conversations(folder, USER_EMAIL)
    print_unread_conversations(unread_conversations)

    unread_items = get_unread_emails(folder)
    print_unread_emails(unread_items)

    results = search_email_by_subject(folder, "invoice")
    print_search_results(results)