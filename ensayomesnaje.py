from twilio.rest import Client 
 
account_sid = 'AC20c89fabb23bc1e94bb8547cbd710ac5' 
auth_token = 'f559bc59b29dcb60aa2d68fd3a21b177' 
client = Client(account_sid, auth_token) 
 
message = client.messages.create( 
                              from_='whatsapp:+14155238886',  
                              body='Your appointment is coming up on July 21 at 3PM',      
                              to='whatsapp:+573006294046' 
                          ) 
 
print(message.sid)