from tkinter import *
import pyodbc

try:
    conn = pyodbc.connect(
        r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\user\Desktop\ICT\students.py;')
    cursor = conn.cursor()
    print("Connected To Database")
except pyodbc.Error as e:
    print("Error in Connection", e)
#this statement corrects the error
#The 'import' statements bring in the required libraries, including Tkinter for GUI and pyodbc for database connectivity.
#The 'try' block attempts to establish a connection to a Microsoft Access database using pyodbc.
#The connection string specifies the driver and the path to the Access database file.
#If the connection is successful, a cursor is created, which is used to execute SQL commands.
#If there is an error in the connection, it is caught, and an error message is printed.

# ClearRecord function
global index
index = None


def ClearRecord():
    global index
    index = None
    sNameValue.set("")
    fNameValue.set("")
    cnicValue.set("")
    cityValue.set("")
    marksValue.set("")
    message.config(text="Record cleared Successfully!", foreground="green")


# FirstRecord function
def FirstRecord():
    try:
        global index
        import pyodbc
        path = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\user\Desktop\ICT\DATABASE H.accdb;'
        connection = pyodbc.connect(path)
        cursor = connection.cursor()
        cursor.execute('SELECT * FROM students')
        result = cursor.fetchall()
        print(result)
        if (result):
            index = 0
            sNameValue.set(result[index][0])
            fNameValue.set(result[index][1])
            cnicValue.set(result[index][2])
            cityValue.set(result[index][3])
            marksValue.set(result[index][4])
            message.config(text="First Record Found!", foreground="green")
        else:
            message.config(text="No Record Found!", foreground="red")
    except pyodbc.Error as e:
        print(f"Could not connect to the database: {e}")
    print("First Record")


# NextRecord function
def NextRecord():
    try:
        global index
        import pyodbc
        path = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\user\Desktop\ICT\DATABASE H.accdb;'
        connection = pyodbc.connect(path)
        cursor = connection.cursor()
        cursor.execute('SELECT * FROM students')
        result = cursor.fetchall()
        if (index is None):
            index = 0
        else:
            index = index + 1
        if (index == len(result)):
            message.config(text="No Further Records Found!", foreground="red")
            index = index - 1
        else:
            if (result[index]):
                sNameValue.set(result[index][0])
                fNameValue.set(result[index][1])
                cnicValue.set(result[index][2])
                cityValue.set(result[index][3])
                marksValue.set(result[index][4])
                message.config(text="Next Record Found!", foreground="green")
            else:
                message.config(text="No Records Found!", foreground="red")
    except pyodbc.Error as e:
        print(f"Could not connect to the database: {e}")
    print("Next Record")


# PreviousRecord function
def PreviousRecord():
    try:
        global index
        import pyodbc
        path = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\user\Desktop\ICT\DATABASE H.accdb;'
        connection = pyodbc.connect(path)
        cursor = connection.cursor()
        cursor.execute('SELECT * FROM students')
        result = cursor.fetchall()
        if (index is None):
            index = (len(result) - 1)
        else:
            index = index - 1
        if (index == -1):
            message.config(text="No Further Records Found!", foreground="red")
            index = index + 1
        else:
            if (result[index]):
                sNameValue.set(result[index][0])
                fNameValue.set(result[index][1])
                cnicValue.set(result[index][2])
                cityValue.set(result[index][3])
                marksValue.set(result[index][4])
                message.config(text="Previous Record Found!", foreground="green")
            else:
                message.config(text="No Records Found!", foreground="red")
    except pyodbc.Error as e:
        print(f"Could not connect to the database: {e}")
    print("Previous Record")


# LastRecord function
def LastRecord():
    try:
        global index
        import pyodbc
        path = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\user\Desktop\ICT\DATABASE H.accdb;'
        connection = pyodbc.connect(path)
        cursor = connection.cursor()
        cursor.execute('SELECT * FROM students')
        result = cursor.fetchall()
        index = len(result) - 1
        if (index):
            sNameValue.set(result[index][0])
            fNameValue.set(result[index][1])
            cnicValue.set(result[index][2])
            cityValue.set(result[index][3])
            marksValue.set(result[index][4])
            message.config(text="Last Record Found!", foreground="green")
        else:
            message.config(text="No Record Found!", foreground="red")
    except pyodbc.Error as e:
        print(f"Could not connect to the database: {e}")
    print("Last Record")


# InsertRecord function
def InsertRecord():
    if (cnicValue.get() == ""):
        message.config(text="Please Enter CNIC.", foreground="red")
    else:
        try:
            import pyodbc
            path = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\user\Desktop\ICT\DATABASE H.accdb;'
            connection = pyodbc.connect(path)
            cursor = connection.cursor()
            sname = sNameValue.get()
            fname = fNameValue.get()
            cnic = cnicValue.get()
            city = cityValue.get()
            marks = marksValue.get()
            cursor.execute('SELECT cnic FROM students WHERE cnic = ?', cnic)
            result = cursor.fetchone()
            if (result):
                message.config(text="CNIC is not unique.", foreground="red")
            else:
                user = (
                    (sname, fname, cnic, city, marks),
                )
                message.config(text="Record added successfully!", foreground="green")
                cursor.executemany(
                    'INSERT INTO students (sname, fname, cnic, city, marks) VALUES (?,?,?,?,?)', user)
                connection.commit()
        except pyodbc.Error as e:
                print(f"Could not connect to the database: {e}")


print("Insert Record")

#The InsertRecord function is triggered when the user wants to add a new record to the database.
#It checks if the CNIC value is provided; if not, it displays an error message.
#It then connects to the database and checks if the CNIC is unique by executing a SELECT query.
#If the CNIC is not unique, an error message is displayed. Otherwise, the record is inserted into the database.
#The executemany method is used to insert a single record into the database.


# UpdateRecord function
def UpdateRecord():
    if (cnicValue.get() == ""):
        message.config(text="Please Enter CNIC.", foreground="red")
    else:
        try:
            import pyodbc
            path = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\user\Desktop\ICT\DATABASE H.accdb;'
            connection = pyodbc.connect(path)
            cnic = cnicEntery.get()
            cursor = connection.cursor()
            cursor.execute('SELECT CNIC FROM students WHERE CNIC = ?', cnic)
            result = cursor.fetchone()
            if (result):
                sname = sNameEntery.get()
                fname = fNameEntery.get()
                cnic = cnicEntery.get()
                city = cityEntery.get()
                marks = marksEntery.get()
                cursor.execute(
                    'UPDATE students SET sname = ?, fname = ?, cnic = ?, city = ?, marks = ? WHERE cnic = ?',
                    (sname, fname, cnic, city, marks, cnic)  # Wrap parameters in a tuple
                )
                connection.commit()
                message.config(text="Record Updated Successfully!", foreground="green")
                sNameValue.set("")
                fNameValue.set("")
                cnicValue.set("")
                cityValue.set("")
                marksValue.set("")
            else:
                message.config(text="No Record Found!", foreground="red")
        except pyodbc.Error as e:
            print(f"Could not connect to the database: {e}")
    print("Update Record")

#The UpdateRecord function is called when the user wants to modify an existing record in the database.
#It checks if the CNIC value is provided; if not, it displays an error message.
#It then connects to the database and checks if a record with the given CNIC exists.
#If the record exists, it updates the record with the values provided in the Entry widgets.
#After updating, it clears the Entry widgets and displays a success message.


# DeleteRecord function
def DeleteRecord():
    if (cnicValue.get() == ""):
        message.config(text="Please Enter CNIC.", foreground="red")
    else:
        try:
            import pyodbc
            path = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\user\Desktop\ICT\DATABASE H.accdb;'
            connection = pyodbc.connect(path)
            cnic = cnicEntery.get()
            cursor = connection.cursor()
            cursor.execute('SELECT CNIC FROM students WHERE CNIC = ?', cnic)
            result = cursor.fetchone()
            if (result):
                cursor.execute('DELETE FROM students WHERE CNIC = ?', cnic)
                connection.commit()
                message.config(text="Record Deleted Successfully!", foreground="green")
                sNameValue.set("")
                fNameValue.set("")
                cnicValue.set("")
                cityValue.set("")
                marksValue.set("")
            else:
                message.config(text="No Record Found!", foreground="red")
        except pyodbc.Error as e:
            print(f"Could not connect to the database: {e}")
    print("Delete Record")

#The DeleteRecord function is called when the user wants to remove a record from the database.
#It checks if the CNIC value is provided; if not, it displays an error message.
#It then connects to the database, checks if a record with the given CNIC exists, and deletes it if found.
#After deleting, it clears the Entry widgets and displays a success message.


# SearchRecord function
def SearchRecord():
    if (cnicValue.get() == ""):
        message.config(text="Please Enter CNIC.", foreground="red")
    else:
        try:
            global index
            path = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\user\Desktop\ICT\DATABASE H.accdb;'
            connection = pyodbc.connect(path)
            cnic = cnicEntery.get()
            cursor = connection.cursor()
            cursor.execute('SELECT CNIC FROM students WHERE CNIC = ?', cnic)
            found = cursor.fetchone()
            cursor.execute('SELECT * FROM students')
            data = cursor.fetchall()
            for i in range(len(data)):
                for j in data[i]:
                    if (j == cnic):
                        index = i
            if (found):
                result = data[index]
                message.config(text="Record Found!", foreground="green")
                sNameValue.set(result[0])
                fNameValue.set(result[1])
                cityValue.set(result[3])
                marksValue.set(result[4])
            else:
                message.config(text="No Record Found!", foreground="red")
        except pyodbc.Error as e:
            print(f"Could not connect to the database: {e}")
    print("Search Record")



#The SearchRecord function is called when the user wants to find a record based on the CNIC.
#It checks if the CNIC value is provided; if not, it displays an error message.
#It connects to the database, checks if a record with the given CNIC exists, and displays it if found.
#The global variable index is used to keep track of the record's index for navigation purposes.


# Design the Student Database Form
root = Tk()
root.geometry("600x400")

#
Label(root, text="Student Database Form", font="Arial 12 bold", foreground='blue').grid(row=0, column=0)
message = Label(root, text="Message Will Appear Here!", foreground='red')
sname = Label(root, text='Student Name', font="ar 10 bold")
fname = Label(root, text='Father Name', font="ar 10 bold")
cnic = Label(root, text='CNIC# (P.Key)', font="ar 10 bold")
search = Label(root, text='Search Record', font="ar 10 bold")
city = Label(root, text='City', font="ar 10 bold")
marks = Label(root, text='Marks', font="ar 10 bold")

message.grid(row=0, column=1)
sname.grid(row=2, column=0)
fname.grid(row=3, column=0)
cnic.grid(row=4, column=0)
search.grid(row=4, column=2)
city.grid(row=5, column=0)
marks.grid(row=6, column=0)

sNameValue = StringVar()
fNameValue = StringVar()
cnicValue = StringVar()
cityValue = StringVar()
marksValue = IntVar()

sNameEntery = Entry(root, textvariable=sNameValue, width='30', font='ar 12 bold')
fNameEntery = Entry(root, textvariable=fNameValue, width='30', font='ar 12 bold')
cnicEntery = Entry(root, textvariable=cnicValue, width='30', font='ar 12 bold')
cityEntery = Entry(root, textvariable=cityValue, width='30', font='ar 12 bold')
marksEntery = Entry(root, textvariable=marksValue, width='30', font='ar 12 bold')

sNameEntery.grid(row=2, column=1, pady=15)
fNameEntery.grid(row=3, column=1, pady=15)
cnicEntery.grid(row=4, column=1, pady=15)
cityEntery.grid(row=5, column=1, pady=15)
marksEntery.grid(row=6, column=1, pady=15)

Button(text="CLEAR", command=ClearRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=7, column=0)
Button(text="FIRST", command=FirstRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=7, column=1)
Button(text="NEXT", command=NextRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=7, column=2)
Button(text="PREVIOUS", command=PreviousRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=9,
                                                                                                              column=0)
Button(text="LAST", command=LastRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=9, column=1)
Button(text="INSERT", command=InsertRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=9,
                                                                                                          column=2)
Button(text="UPDATE", command=UpdateRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=11,
                                                                                                          column=0)
Button(text="DELETE", command=DeleteRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=11,
                                                                                                          column=1)
Button(text="SEARCH", command=SearchRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=11,
                                                                                                          column=2)

root.mainloop()
