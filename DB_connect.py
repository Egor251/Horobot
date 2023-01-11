import sqlite3

def select(sql):
    conn = sqlite3.connect("mydatabase.db")  # или :memory: чтобы сохранить в RAM
    cursor = conn.cursor()
    cursor.execute(sql)
    result = cursor.fetchall()
    print (result)

sql = "SELECT * FROM programs where teacher='Тираспольская Екатерина Ильинична'"
select(sql)
a= 1
