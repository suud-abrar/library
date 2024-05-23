import sqlite3


conn = sqlite3.connect('database.db')
print('connected successfully')
conn.execute('CREATE TABLE signup (username TEXT, password TEXT)')
print('create successfully')

conn.close()
