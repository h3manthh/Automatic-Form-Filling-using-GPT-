import bcrypt

password = "qwerty12345"

# Hash the password
hashed_password = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())

print(hashed_password.decode('utf-8'))
