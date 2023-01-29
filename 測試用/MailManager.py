import email


def parse_msg(filepath):
    with open(filepath, 'rb') as f:
        message = email.message_from_binary_file(f)
    headers = {}
    for name, value in message.items():
        headers[name] = value
    subject = message['subject']
    sender = message['from']
    recipient = message['to']
    content = message.get_payload()
    return {
        'headers': headers,
        'subject': subject,
        'sender': sender,
        'recipient': recipient,
        'content': content
    }


print(parse_msg("./SimpleMessage.msg")["content"])
