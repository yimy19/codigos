import cryptocode
import base64

ensayo=base64.b64encode('ensayo')
print(ensayo)

str_encoded = cryptocode.encrypt('ensayo', 'reditos')
print(str_encoded)
## And then to decode it:
str_decoded = cryptocode.decrypt(str_encoded, 'reditos')
print(str_decoded)
