from twilio.rest import Client

# Your Account SID pip3
account_sid = "AC82e1ac0645f026fb9a1447739e985c58"
# Your Auth Token from twilio.com/console
auth_token  = "fd4c2dc5c36971b7775edda7bb38ce06"

client = Client(account_sid, auth_token)

message = client.messages.create(
    to="+573006294046", 
    from_="+19895755564",
    body="Hello from Python!")

print(message.sid)