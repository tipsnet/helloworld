import imaplib
import email
import pickle
import socks
import socket


import ssl

class SSLContextPickler(pickle.Pickler):
    def persistent_id(self, obj):
        if isinstance(obj, ssl.SSLContext):
            return 'SSLContext'
        return None

class SSLContextUnpickler(pickle.Unpickler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.ssl_context = None

    def persistent_load(self, pid):
        if pid == 'SSLContext':
            if self.ssl_context is None:
                self.ssl_context = ssl.create_default_context()
            return self.ssl_context
        raise pickle.UnpicklingError(f"unsupported persistent id: {pid}")



def login(email, password):
    # create an SSL context
    ssl_context = ssl.create_default_context()

    # imap_server = "imap-mail.outlook.com"
    imap_server = "outlook.office365.com"

    
    # imap_server_conn = imaplib.IMAP4_SSL(imap_server, 993, context=ssl_context)
    imap_server_conn = imaplib.IMAP4_SSL(imap_server, 993, ssl_context=ssl_context)
    
    imap_server_conn.login(email, password)
    return imap_server_conn, ssl_context

def save_session(session_file, imap_server_conn, ssl_context):
    with open(session_file, "wb") as f:
        pickler = SSLContextPickler(f)
        pickler.dump((imap_server_conn, ssl_context))

def load_session(session_file):
    with open(session_file, "rb") as f:
        unpickler = SSLContextUnpickler(f)
        imap_server_conn, ssl_context = unpickler.load()
    return imap_server_conn, ssl_context



def save_session_v3(session_file, imap_server_conn, ssl_context):
    with open(session_file, "wb") as f:
        pickler = pickle.Pickler(f)
        pickler.dump((imap_server_conn.host, imap_server_conn.port, imap_server_conn.state))
        pickler.dump(ssl_context.__dict__)

def save_session_v4(session_file, imap_server_conn):
    with open(session_file, "wb") as f:
        pickle.dump(imap_server_conn, f)
# 171.242.41.237:17747
# proxy_ip = '171.242.41.237'
# proxy_port = 17747

# # Set up the SOCKS5 proxy
# socks.set_default_proxy(socks.SOCKS5, proxy_ip, proxy_port)
# socket.socket = socks.socksocket

# Connect to the IMAP server
imap_server = imaplib.IMAP4_SSL("outlook.office365.com", 993)

# Log in to the server
username = "olgaamber9ze@hotmail.com"
password = "vAh4ewd4fu"
# imap_server.login(username, password)

imap_server, ssl_context = login(username, password)

# save session file
session_file = 'sessions/' + username + '.pickle'
save_session_v4(session_file, imap_server)

# Read emails from the inbox

# Select the inbox
imap_server.select("inbox")

# Search for all emails in the inbox
result, data = imap_server.search(None, "ALL")

# print(result)
# print(data)


# Loop through all the emails and print out the subject and sender
for num in data[0].split():
    # Get the email data and parse it into an email object
    result, email_data = imap_server.fetch(num, "(RFC822)")
    email_msg = email.message_from_bytes(email_data[0][1])

    subject = email_msg.get('Subject')
    sender = email_msg.get('From')
    recipient = email_msg.get('To')
    date = email_msg.get('Date')
    message_id = email_msg.get('Message-ID')
    in_reply_to = email_msg.get('In-Reply-To')
    references = email_msg.get('References')

    # Print out the subject and sender of the email
    print("sender:", sender)
    print("Subject:", email_msg["Subject"])
    print("From:", email.utils.parseaddr(email_msg["From"]))
    print("date:", date)
    print("message_id:", message_id)
    print("in_reply_to:", in_reply_to)
    print("references:", references)
    print(result)
    
    

    # Print out the body of the email
    # for part in email_message.walk():
    #     print(part.get_content_type(), part.get_content_maintype(), part.get_content_subtype())
    #     print('---------')
    #     if part.get_content_type() == "text/plain":
    #         body = part.get_payload(decode=True)
    #         print(body)
    #     print('.................')

    email_body = None
    # get the plain text body of the email
    if email_msg.is_multipart():
        for part in email_msg.walk():
            content_type = part.get_content_type()
            if content_type == 'text/plain':
                email_body = part.get_payload(decode=True).decode('utf-8')
    else:
        email_body = email_msg.get_payload(decode=True).decode('utf-8')

    # print the email body
    print(email_body)

    print('----------------------------------------------------------------------')

# Log out of the server
imap_server.logout()
