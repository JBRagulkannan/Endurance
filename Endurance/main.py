# import datetime
#
# import _mysql_connector
# import mysql.connector
# from flask import Flask,render_template,request
# app = Flask(__name__)
#
# con=mysql.connector.Connect(host="localhost",user="root",password="12345678",database="Endurance_db")
#
# cur_datetime = None
#
# def get_cur_datetime():
# 	cur_datetime = datetime.datetime.now()
# 	cur_datetime = cur_datetime.date()
# 	cur_datetime = cur_datetime.today()
# 	return cur_datetime
#
#
# def insert(initial, final):
# 	cursor = con.cursor()
# 	query = "INSERT INTO users (initial, final) VALUES (%s, %s)"
# 	cursor.execute(query, (initial, final))
# 	con.commit()
# 	cursor.close()
# if con:
# 	print("connect")
# 	cur_datetime = get_cur_datetime()
# else:
# 	print("not connect")
#
#
#
# @app.route('/')
# @app.route('/home')
# def home():
# 	return render_template('home page.html')
#
# @app.route("/data",methods=['POST','GET'])
# def endurance():
# 	if request.method=='POST':
# 		i=request.form.get('Initialvalue')
# 		f=request.form.get('Finalvalue')
# 		print(i,f)
# 		if i is not None:
# 			insert(i,None)
# 		# insert(i,f)
# 		elif f is not None:
# 			insert(None,f)
# 		print(cur_datetime)
# 		return render_template('home page.html',Initialvalue=i,Finalvalue=f)
#
# if __name__ == '__main__':
# 	app.run(port=8002)
#

import datetime
import serial
import serial.tools.list_ports
import mysql.connector
import _tkinter
from tkinter import *

from flask import Flask, render_template, request, session




app = Flask(__name__)
app.secret_key = b'_5#y2L"F4Q8z\n\xec]/'

con = mysql.connector.Connect(host="localhost", user="root", password="12345678", database="Endurance_db")

cur_datetime = None

currtime = None
com_port = None
def insert(initial,final,curr_date,isInitial):
    cursor = con.cursor()
    data = {
        'initial':initial,
        'final':final,
        'curr_date':curr_date
    }
    if isInitial:
        query = "insert into users(initial,final,curr_date)VALUES (%(initial)s, %(final)s, %(curr_date)s) ON DUPLICATE KEY UPDATE initial=Values(initial),curr_date = Values(curr_date)"
    else:
        query = "insert into users(initial,final,curr_date)VALUES (%(initial)s, %(final)s, %(curr_date)s) ON DUPLICATE KEY UPDATE final=Values(final),curr_date = Values(curr_date)"
    cursor.execute(query,data)
    con.commit()
    cursor.close()

def curr_dt():
    global currtime
    currtime = datetime.datetime.now().strftime('%Y-%m-%d')
    # '%H:%M:%S'
    global cur_datetime
    cur_datetime = datetime.datetime.now().strftime('%Y-%m-%d')
    print(cur_datetime)


def comport_selection():
    global com_port
    com_port = []
    com_ports = serial.tools.list_ports.comports()
    for port in com_ports:
        # print(port)
        # print(type(port))
        com_port.append(port)



@app.route('/')
@app.route('/home')
def home():
    curr_dt()
    comport_selection()
    print(cur_datetime)
    last_selected_value = session.get('last_selected_value', None)
    return render_template('home page.html', current_datetime=currtime, Initialvalue=None, Finalvalue=None,
                           com_port=com_port,last_selected_value=last_selected_value)


@app.route("/data", methods=['POST', 'GET'])
def endurance():
    if request.method == 'POST':
        curr_dt()
        comport_selection()
        i = request.form.get('Initialvalue')
        f = request.form.get('Finalvalue')
        get_comport = request.form.get('comport')
        session.setdefault('last_selected_value',get_comport)
        print(get_comport)
        if 'initialSubmit' in request.form:
            insert(i, None,cur_datetime,True)
            return render_template('home page.html',current_datetime=currtime,Initialvalue=i, Finalvalue=None, com_port=com_port,last_selected_value=get_comport)
        elif 'finalSubmit' in request.form:
            insert(None, f,cur_datetime,False)
            return render_template('home page.html', current_datetime=currtime, Initialvalue=None, Finalvalue=f, com_port=com_port,last_selected_value=get_comport)
    return render_template('home page.html', current_datetime=currtime, Initialvalue=None, Finalvalue=None,com_port=com_port,last_selected_value=get_comport)

if __name__ == '__main__':
    app.run(port=8002)
