import base64

with open('icons8-table-32.png', 'rb') as open_icon:
    b64str = base64.b64encode(open_icon.read())
    print(b64str)