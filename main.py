import ctypes
import time
import functions
from win32gui import GetWindowText, GetForegroundWindow
import datetime
import win32
import win32com.client
import csv
import shutil
import re

#Script to log window activity to be used to enter timesheets weekly.
def timesheet():
    database = r"C:\Users\tduffett\PycharmProjects\Timesheet\timesheet.db"
    # SQL query to get programs + times
    conn = functions.create_connection(database)

    with conn:
        # update foregrounds table
        sql = """SELECT * FROM foregrounds WHERE time > (?)"""
        cur = conn.cursor()
        cur.execute(sql,('20240301104752',))
        conn.commit()
        return cur.fetchall()
def log():
        #Set database to store values
        database = r"C:\Users\tduffett\PycharmProjects\Timesheet\timesheet.db"
        foreground_old = None
        while True:

            #Get Outlook MAPI
            #mapi = functions.get_outlook()

            #Use recurssion to iterate through folders
            #functions.outlook_search(mapi, '\\')


            # Add foreground window and time into database
            ct = datetime.datetime.now().strftime("%Y%m%d%H%M%S")

            foreground_new = GetWindowText(GetForegroundWindow())
            if foreground_new != foreground_old:
                functions.add_foregrounds(database, foreground_new, ct)
            foreground_old = foreground_new
            #Get list of all open windows
            #print(list(filter(None,functions.GetAllWindows())))

            #functions.add_programs(database,list(filter(None,functions.GetAllWindows())),ct)

            #Scan for changes to bluebeam files

            #Save Copy of .dat file for comparison as RdbRecentFiles.dat?

            #If no .dat file already in Timesheet folder
            try:
                Bluebeam_old = ""
                Bluebeam_old.join(i.strip() for i in open(r"BluebeamLog.dat",encoding='latin-1').readlines())
                Bluebeam_new = [i.strip() for i in
                              open(r"C:\Users\tduffett\AppData\Roaming\Bluebeam Software\Revu\20\RdbRecentFiles.dat",
                                   encoding='latin-1').readlines()]
                #Check if new and old are the same length, if not then all new lines are new locations visited. There may still
                #be visted locations so do not use else
                locations = []
                if len(Bluebeam_new) != len(Bluebeam_old):
                    for i in range(len(Bluebeam_old),len(Bluebeam_new)):
                        #Store all
                        #print(Bluebeam_new[i])
                        locations += (re.findall(':\\\.*?pdf', Bluebeam_new[i]))

                #Read line for line and look for changes:
                for i in range(len(Bluebeam_old)):
                    if Bluebeam_new[i] != Bluebeam_old[i]:
                        print(Bluebeam_new[i])
                        locations += re.findall(':\\\.*?pdf', Bluebeam_new[i])

                #print(locations)


                    #If not the same length then need to add every pdf in the new lines.



                # Continue with comparison

                # read each line
                # compare with previous file
                # if change then locate where change was and record file location preceeding change location.
                # if change is at end of document, then this is a new document and should also read file location of new file.

                # If change, from back of line find last ".pdf" and read to "C:\"



            except:
                print("ERROR")
                #shutil.copy2(r"C:\Users\tduffett\AppData\Roaming\Bluebeam Software\Revu\20\RdbRecentFiles.dat",r"C:\Users\tduffett\PycharmProjects\Timesheet\BluebeamLog.dat")

                # then copy new one from bluebeam folder



            #datContent = [i.strip() for i in open(r"C:\Users\tduffett\AppData\Roaming\Bluebeam Software\Revu\20\RdbRecentFiles.dat",encoding='cp437').readlines()]
            #print(datContent[2])


            #Repeat indefinitely while PC awake
            time.sleep(1)


            #C:\Users\tduffett\AppData\Roaming\Bluebeam Software\Revu\20\RdbRecentFiles.dat is where bluebeam recent files are located




if __name__ == "__main__":
    log()
