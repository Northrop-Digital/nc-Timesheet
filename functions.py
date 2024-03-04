import ctypes
import sqlite3
from sqlite3 import Error
from win32gui import GetWindowText, GetForegroundWindow
import win32com.client

EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible

titles = []
def foreach_window(hwnd, lParam):
    if IsWindowVisible(hwnd):
        length = GetWindowTextLength(hwnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        GetWindowText(hwnd, buff, length + 1)
        titles.append((buff.value))
    return True
def GetAllWindows():
    EnumWindows(EnumWindowsProc(foreach_window), 0)
    return titles

def create_connection(db_file):
    """ create a database connection to the SQLite database
        specified by db_file
    :param db_file: database file
    :return: Connection object or None
    """
    conn = None
    try:
        conn = sqlite3.connect(db_file)
    except Error as e:
        print(e)

    return conn

def add_foregrounds(database,foreground,time):
    conn = create_connection(database)

    with conn:
        # update foregrounds table
        sql = """INSERT INTO foregrounds(foreground,time) VALUES (?,?)"""
        cur = conn.cursor()
        cur.execute(sql,(foreground,str(time)))
        conn.commit()
        return cur.lastrowid

def add_programs(database,programs,time):
    conn = create_connection(database)

    with conn:
       # update programs table
        for i in range(len(programs)):
            sql = """INSERT INTO programs(program,time) VALUES (?,?)"""
            cur = conn.cursor()
            cur.execute(sql,(programs[i],str(time)))
            conn.commit()
        return cur.lastrowid

def outlook_search(Folder, path):
    path += str(Folder) + "\\"
    for i in Folder.Folders:
        outlook_search(i, path)
    else:
        for message in Folder.Items:
            print(path, message.ReceivedTime, message.Subject)
            pass
        return

def get_outlook():
    outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI").Folders['tduffett@northrop.com.au'].Folders['Inbox'].Folders['Projects']
    return mapi