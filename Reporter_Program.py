"""
  *******************************************************************************************************************
  *                                              Reporter_Program                                                   *
  *******************************************************************************************************************
  *                                                                                                                 *
  *  This program works by checking a directory, the default directory is wherever you put the program.             *
  *  When there is an update action performed on the destination directory, it grabs that file and, if the          *
  *  updated file is a csv filetype, it parses the file for the last entry. It then                                 *
  *  pulls all of the information in that final line and inserts it into the appropriate SQL server table. If       *
  *  anything goes wrong, it logs it in the locally created "Failed" folder. All properly inserted SQL statements   *
  *  are timestamped and put in the "Archive" folder.                                                               *
  *                                                                                                                 *
  *  Every good SQL commit prints "CSV Committed" to the command window, mostly for a quick visual debugging.       *
  *                                                                                                                 *
  *******************************************************************************************************************
"""

from __future__ import generators
from distutils.core import setup
import os
import sys
import datetime
import queue as Queue
import threading
import time
import win32file
import win32con
import pyodbc
import csv
from collections import deque
import string
import traceback  
from watchdog.observers import Observer  
from watchdog.events import PatternMatchingEventHandler 
from functools import partial

# If the files for logging success & failure don't exist, make them.
if not os.path.exists('Archive'):
    os.mkdir("Archive")
if not os.path.exists('Failed'):
    os.mkdir("Failed")

ReportPath = os.path.dirname(os.path.realpath(__file__))
#If label making goes well, push stuff here.
archive_path = str(ReportPath) + "\\Report Archive"
#If we can't print a label, or can print but with errors, put those here.
fail_path = str(ReportPath) + "\\Failed Reports"
# Global variable to hold onto the changing values of the old Serial Number.
old_SN = 'Start'

# This check is what helps to compare UIDs against each other. In this case 
# we are comparing the Serial_Numbers against each other. This is necessary
# because this specific test machine doesn't change it's generated Serial_Number
# until the test is passed, so we only want to use the last Serial_Number.
def CheckSN(Serial_Number):
  # If the SN is the same as the last one, discard.
  #print ('Started CheckSN.') #debug
  global old_SN
  #print ('passed CheckSN Print Statements.') #debug
  if old_SN == Serial_Number:
    print('Duplicate SN, ignoring it and monitoring again.')
    return ('Duplicate')
  # If this is the first time the program catches a SN, discard.
  #print ('Passed first hurdle.')
  elif old_SN == 'Start':
    print ('First SN. Storing it and monitoring again.')
    old_SN = Serial_Number
    return ('Duplicate')
  # When the program encounters a new SN, it pushes it to PyReporter_Program
  #print ('Passed second hurdle.')
  elif Serial_Number != old_SN:
    print ('New SN.')
    # This is the assign statement to move the current Serial_Number into the global variable.
    old_SN = Serial_Number
    return(Serial_Number)
  else:
    print('No catches in CheckSN.')
    return ('Duplicate.')

# This function parses the .csv f/ile and returns the last row as, you guessed it, "lastrow"
def get_last_row(csv_filename):
    with open(csv_filename, 'r') as f:
        try:
            lastrow = deque(csv.reader(f), 1)[0]
        # This should throw an error if the file is empty, which should continue the watcher.
        except IndexError:
            lastrow = None
        #print (lastrow) # debug
        return (lastrow)
 
# This is the main function. It gets passed the name of the file from the watcher class and then pushes it
# to sql.
def PyReporter_Program(Reporter_ProgramFile):

  try:
    # Instantiates get_last_row with the file as an argument. Assigns it to SQLQuery.
    SQLQuery = get_last_row(Reporter_ProgramFile) 
    # Multiple assign statement to pull variables from SQLQuery
    Part_Number, SlipTable_Description, Serial_Number, Inlet_Clearance, Center_Clearance, Discharge_Clearance, Front_Clearance, Back_Clearance, Float_Clearance, Gear_End_Clearance_Drive_N, Gear_End_Clearance_Drive, Drive_End_Clearance_Drive_N, Drive_End_Clearance_Drive, PSI_Slip_RPM, SlipTime = SQLQuery
    # This particular time variable wasn't particularly important (it ended up losing less than a second) 
    # and could be substituted for a datetime stamp for the datetime type in SQL
    SlipTime = 'getdate()'
  except:
    # Except statements work to error handle each compartmentalized part of the process.
    PyLogging = open('' + fail_path + '\\PyLogging.txt', 'a+')
    PyLogging.write('\n' + 'Timestamp: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()) + "\n    get_last_row() failed to parse the csv file at all.")
    PyLogging.close()
    print('''Python couldn't find your newest Slip Test file. Tell IT to "Check if the Slip text files are stored at ''' + fail_path + ''', and if so, to check on the Pywatcher.py code block in Eli's 'Nerd Stuff' folder.''')

  try:
    # This block runs a check for duplicate serial numbers and breaks out of PyReporter_Program if the SN is a duplicate.
    #print ('Started Check.') # debug
    PNCheck =  CheckSN(Serial_Number)
    if PNCheck == 'Duplicate':
      return
    else:
      PNCheck = Part_Number
      BLUID = '0000'
  except:
    #Except statements work to error handle each compartmentalized part of the process.
    PyLogging = open('' + fail_path + '\\PyLogging.txt', 'a+')
    PyLogging.write('\n' + 'Timestamp: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()) + "\n    CheckSN() failed. This could have to do with the uniqueness of the Serial Number. There is no regex.")
    PyLogging.close()
    print('''Serial_Number duplication check failed. Tell IT to look at the error log at: ''' + fail_path + '''\PyLogging.txt''')
    
  try:
    #TODO Hide this in a private function & make a new user (this one is also Mary Paris' login).
    #SQL contact and insert statements
    server = '*******'
    database = '*******'
    username = '*******'
    password = '*******'
    driver = '{ODBC Driver 13 for SQL Server}'
    table = 'dbo.*******'
    cnxn = pyodbc.connect('DRIVER=' + driver + ';SERVER=' + server + ';PORT=1443;DATABASE=' + database + ';UID=' + username + ';PWD='+ password)
    cursor = cnxn.cursor()
  except:
    #Except statements work to error handle each compartmentalized part of the process.
    #TODO This needs to be a popup.
    PyLogging = open('' + fail_path + '\\PyLogging.txt', 'a+')
    PyLogging.write('\n' + 'Timestamp: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()) + "\n    pyodbc.connect module failed to connect to USCEVAPP006.")
    PyLogging.close()
    print ("Python couldn't establish a connection to the databse. Tell IT to look at the error log at: " + fail_path + "\PyLogging.txt")

  try:
    #This is the insert statement for SQL from PyODBC, and is contingent on PYODBC drivers.
    cursor.execute("""INSERT INTO """ + table + """ (
          Part_Number, Description, Serial_Number, Inlet_Clearance, Center_Clearance, Discharge_Clearance, Front_Clearance, Back_Clearance, Float_Clearance, Gear_End_Clearance_Drive_N, Gear_End_Clearance_Drive, Drive_End_Clearance_Drive_N, Drive_End_Clearance_Drive, PSI_Slip_RPM, DateTime, BL_UID)
           VALUES ('""" + Part_Number + """', '""" + SlipTable_Description + """', RTRIM('""" + Serial_Number + """'), '""" + Inlet_Clearance + """', '""" + Center_Clearance + """', '""" + Discharge_Clearance + """', '""" + Front_Clearance + """', '""" + Back_Clearance + """', '""" + Float_Clearance + """', '""" + Gear_End_Clearance_Drive_N + """', '""" + Gear_End_Clearance_Drive + """', '""" + Drive_End_Clearance_Drive_N + """', '""" + Drive_End_Clearance_Drive + """', '""" + PSI_Slip_RPM + """', '""" + '{:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()) + """', '""" + BLUID + """')""")
    cnxn.commit()
    PyLogging = open('' + archive_path + '\\PyLogging.txt', 'a+')
    PyLogging.write('\n' + 'Timestamp: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()))
    PyLogging.close()
    print ('CSV has been committed.') # debug
  except:
    #Except statements work to error handle each compartmentalized part of the process.
    #TODO This needs to be a popup.
    PyLogging = open('' + fail_path + '\\PyLogging.txt', 'a+')
    PyLogging.write('\n' + 'Timestamp: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()) + "\n    Insert statement failed for pyodbc module. Commit statement is crucial but will not cause failure. Try looking at SQL table data types, must be varchar of the proper length.")
    PyLogging.close()
    print ("Python couldn't write to the database. Tell IT to look at the error log at: " + fail_path + "\PyLogging.txt")

  cnxn.close()

# This just assigns numbers to the possible actions that the watcher class covers. We will later error handle to only
# worry about the "Updated" action because both create and update actions in windows trigger the Update action.
ACTIONS = {
  1 : "Created",
  2 : "Deleted",
  3 : "Updated",
  4 : "Renamed to something",
  5 : "Renamed from something"
  }

# We've successfully defaulted this to the folder that the Reporter_Program writes to, but it can include any folder with the
# proper arguments and watch all at the same time.
def watch_path (path_to_watch, include_subdirectories=False):

  #Try watching the directory.
  try:
    FILE_LIST_DIRECTORY = 0x0001
    hDir = win32file.CreateFile (
      path_to_watch,
      FILE_LIST_DIRECTORY,
      win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
      None,
      win32con.OPEN_EXISTING,
      win32con.FILE_FLAG_BACKUP_SEMANTICS,
      None
    )
    while 1:
      results = win32file.ReadDirectoryChangesW (
        hDir,
        1024,
        include_subdirectories,
        win32con.FILE_NOTIFY_CHANGE_FILE_NAME | 
         win32con.FILE_NOTIFY_CHANGE_DIR_NAME |
         win32con.FILE_NOTIFY_CHANGE_ATTRIBUTES |
         win32con.FILE_NOTIFY_CHANGE_SIZE |
         win32con.FILE_NOTIFY_CHANGE_LAST_WRITE |
         win32con.FILE_NOTIFY_CHANGE_SECURITY,
        None,
        None
      )
      #When there is an action, it triggers the following process. This then narrows it down to a .csv and further
      #to an update action.
      for action, file in results:
        full_filename = os.path.join (path_to_watch, file)
        filename, file_extension = os.path.splitext(full_filename)
        csv_filename = ['.csv']
        action_list = [3]
        #if its both a .csv and an update action
        if (file_extension in csv_filename) and (action in action_list):
          #obligatory pass on delete
          if not os.path.exists (full_filename):
            pass
          #obligatory pass on directories (this way it watches narrowly enough to only pass the .csv)
          elif os.path.isdir (full_filename):
            pass
          #If this is a .csv, and an update action, and a file, pass the filename to PyReporter_Program
          else: 
            file_type = 'file'
            filename, file_extension = os.path.splitext(full_filename)
            PyReporter_Program(full_filename)
            #PyLogging.close()
        else:
          #print ("type " + file_extension + " Failed .csv & action_list conditional") #debug
          pass
  #Logging in case the Watch_Path function has an uncaught error.
  except:
    PyLogging = open('' + fail_path + '\\PyLogging.txt', 'a+')
    PyLogging.write('\n' + 'Timestamp: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()) + "\n    Watch_Path function failed. Restarting...")
    PyLogging.close()
    
#This is where the rubber hits the road. The Watcher class instantiates the whole process to observe a folder.
class Watcher (threading.Thread):

  def __init__ (self, path_to_watch, results_queue, **kwds):
    threading.Thread.__init__ (self, **kwds)
    self.setDaemon (1)
    self.path_to_watch = path_to_watch
    self.results_queue = results_queue
    self.start ()

  def run (self):
    for result in watch_path (self.path_to_watch):
      self.results_queue.put (result)

if __name__ == '__main__':
  #This is where I set it to default to the Test Data folder, which is where Reporter_Program happens to drop its data.  
  PATH_TO_WATCH = [str(ReportPath)]
  #This is in case there are several paths to watch.
  try: path_to_watch = sys.argv[1].split (",") or PATH_TO_WATCH
  except: path_to_watch = PATH_TO_WATCH
  path_to_watch = [os.path.abspath (p) for p in path_to_watch]

  #confirmation that the program is indeed working.
  print ("Watching %s at %s" % (", ".join (path_to_watch), time.asctime ()))
  files_changed = Queue.Queue ()
  
  #an iterator to parse the changes in the file in order to pass them to the Watcher function
  for p in path_to_watch:
    Watcher (p, files_changed)

  #This loop uses multiple assignment to pull apart anything in the queue and, if nothing, it sleeps.
  while 1:
    try:
      file_type, filename, action = files_changed.get_nowait ()
      #print (file_type, filename, action) #debug
    except Queue.Empty:
      pass
    time.sleep (1)
  
