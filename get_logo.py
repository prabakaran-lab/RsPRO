import base64
data = open('logo.jpg', 'rb').read()
b64 = base64.b64encode(data).decode()
print(f'data:image/jpeg;base64,{b64}', end='')
