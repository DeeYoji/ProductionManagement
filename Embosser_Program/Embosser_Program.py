'''
  **************************************************************************************************************
  *                                          Embosser Program                                                  *
  **************************************************************************************************************
  *                                                                                                            *
  * This program is meant to take input in the form of a barcode from a label put out by the Label_Program.    *
  * It then takes that information and tries to fetch it from the primary database. Failing that, it fetches   *
  * from the secondary database. It then pulls the master file that contains coordinates for the embosser to   *
  * print the letters to specific locations and pushes that information to the embosser. Then, it pushes the   *
  * information back to the SQl databases.                                                                     *
  *                                                                                                            *
  * This particular system generated a serial number on another machine in the middle of the production line.  *
  * We prompt the user for this SN so we can enter it. This is probably unnecessary if your system works       *
  * differently.                                                                                               *
  *                                                                                                            *
  **************************************************************************************************************
'''
import pyodbc
import datetime
import os
import tempfile
import win32api
import win32print
import time

# If we don't have folders for successes and failures, make them in the current folder.
if not os.path.exists('Embosser_Archive'):
    os.mkdir("Embosser_Archive")
if not os.path.exists('Embosser_Failed'):
    os.mkdir("Embosser_Failed")
CurrentDirectory = os.path.dirname(os.path.realpath(__file__))

# This function runs a plate check by the user to make sure that the appropriate plate is embossed.
# It then reads the coordinates in the master file and uses those to emboss the selected fields onto
# the corresponding fields on the plate.
def PlateMaker(PlateType, BOM, PartDescription, SerialNumber, CustomerPartNumber):
  # Notify the user of the type of plate being used.
  while True:
    Plate_Check = input("Please put in a "+PlateType+" plate and press enter to confirm.")
    if Plate_Check:
      continue
    else:
      break
  # This field adds the year as a variable
  Year = '{:%Y}'.format(datetime.datetime.now())
  # We stuff this into filename so we can save it later.
  filename = PlateType.lower().rstrip() + ' Plate'
  #print ("I read: PSONum = "+str(PSONumber)+", Part "+str(BlowerNumber)+" of "+str(TotalBlowers)+" in year "+Year) #debug
  # These blocks narrow the plate selection type so that they can be given different coordinates.
  try:
    if filename == "aa Plate":
      PlateFile = CurrentDirectory+'/Embosser_Archive/Aa_'+SerialNumber+'_'+BOM+'.txt'
      GoesIn = open('Plate_Master_Files/Aa_Master.txt', 'r')
      GoesOut = open(PlateFile, 'w')
      for lines in GoesIn:
        GoesOut.write(lines)
      GoesOut.write ("""\n\n<"""+SerialNumber.rstrip()+"""\n"""+BOM.rstrip()[:11]+"""\n"""+PartDescription.rstrip()[:25]+"""\n"""+Year+""">""")
      win32api.ShellExecute ( 0, "print", PlateFile, '/d:"%s"' % win32print.GetDefaultPrinter (), ".", 0)
      print ("Embossing", SerialNumber, "onto a", PlateType, "plate for", PSONum, ".")

    elif filename == "bb Plate":
      PlateFile = CurrentDirectory+'/Embosser_Archive/Bb_'+SerialNumber+'_'+BOM+'.txt'
      GoesIn = open('Plate_Master_Files/Bb_Master.txt', 'r')
      GoesOut = open(PlateFile, 'w')
      for lines in GoesIn:
        GoesOut.write(lines)
      GoesOut.write ("""\n\n<"""+PartDescription.rstrip().lstrip()[:19]+"""\n"""+SerialNumber.rstrip().lstrip()+"""\n"""+BOM.rstrip().lstrip()+""">""")
      GoesIn.close()
      GoesOut.close()
      win32api.ShellExecute ( 0, "print", PlateFile, '/d:"%s"' % win32print.GetDefaultPrinter (), ".", 0)
      print ("Embossing", SerialNumber, "onto a", PlateType, "plate for", PSONum, ".")

    elif filename == "cc Plate":
      PlateFile = CurrentDirectory+'/Embosser_Archive/Cc_'+SerialNumber+'_'+BOM+'.txt'
      GoesIn = open('Plate_Master_Files/Cc_Master.txt', 'r')
      GoesOut = open(PlateFile, 'w')
      for lines in GoesIn:
        GoesOut.write(lines)
      GoesOut.write ("""\n\n<"""+PartDescription.rstrip().lstrip()[:15]+"""\n"""+SerialNumber.rstrip().lstrip()+"""\n"""+BOM.rstrip().lstrip()+""">""")
      GoesIn.close()
      GoesOut.close()
      win32api.ShellExecute ( 0, "print", PlateFile, '/d:"%s"' % win32print.GetDefaultPrinter (), ".", 0)
      print ("Embossing", SerialNumber, "onto a", PlateType, "plate for", PSONum, ".")

    elif filename == "dd Plate":
      PlateFile = CurrentDirectory+'/Embosser_Archive/Dd_'+SerialNumber+'_'+BOM+'.txt'
      GoesIn = open('Plate_Master_Files/Dd_Master.txt', 'r')
      GoesOut = open(PlateFile, 'w')
      for lines in GoesIn:
        GoesOut.write(lines)
      GoesOut.write ("""\n\n<"""+PartDescription.rstrip()[:14]+"""-"""+CustomerPartNumber.rstrip().lstrip()+"""\n"""+SerialNumber.rstrip()+"""\n"""+BOM.rstrip()+""">""")
      GoesIn.close()
      GoesOut.close()
      win32api.ShellExecute ( 0, "print", PlateFile, '/d:"%s"' % win32print.GetDefaultPrinter (), ".", 0)
      print ("Embossing", SerialNumber, "onto a", PlateType, "plate for", PSONum, ".")

    elif filename == "ee Plate":
      PlateFile = CurrentDirectory+'/Embosser_Archive/Ee_'+SerialNumber+'_'+BOM+'.txt'
      GoesIn = open('Plate_Master_Files/Ee_Master.txt', 'r')
      GoesOut = open(PlateFile, 'w')
      for lines in GoesIn:
        GoesOut.write(lines)
      GoesOut.write ("""\n\n<"""+CustomerPartNumber+"""\n"""+PartDescription.rstrip()[:10]+"""\n"""+SerialNumber.rstrip()+"""\n"""+BOM.rstrip()+""">""")
      GoesIn.close()
      GoesOut.close()
      win32api.ShellExecute ( 0, "print", PlateFile, '/d:"%s"' % win32print.GetDefaultPrinter (), ".", 0)
      print ("Embossing", SerialNumber, "onto a", PlateType, "plate for", PSONum, ".")

  except:
    print ("print block failure.")
    time.sleep(3)

# After the plate prints, we commit changes to the SQL databases so we know that things went well.
def PyPusher(SerialNumber, BOM, Description, PSONum):
  cursor, cnxn = SQL_cnxn()

  try:
    # This is the insert statement for SQL from PyODBC, and is contingent on PYODBC drivers.
    cursor.execute("""INSERT INTO dbo.SerialNumbers ( [SerialNumber], [BOM], [Description], [Year], [DateTime], [PSONum])
      VALUES ('"""+SerialNumber+"""', '"""+BOM+"""', '"""+Description+"""', '"""+str('{:%Y}'.format(datetime.datetime.now()))+"""', GETDATE(), '"""+PSONum+"""')""")
    cnxn.commit()
    PyLogging = open(str(CurrentDirectory)+'\\PyLogging.txt', 'a+')
    PyLogging.write('\n' + 'Timestamp: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()))
    PyLogging.close()
    cnxn.close()
    #print ('Database filled..') # debug
  except:
    PyLogging = open(str(CurrentDirectory) +'\\Embosser_Failed\\PyLogging.txt', 'a+')
    PyLogging.write('\n' + 'Timestamp: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()) + "\n    Insert statement failed for pyodbc module. Commit statement is crucial but will not cause failure. Try looking at SQL table data types, must be varchar of the proper length.")
    PyLogging.close()
    print ("Python couldn't write to the database. Tell IT to look at the error log at: "+str(CurrentDirectory)+"\PyLogging.txt")
    return

  try:
    PlateLog = open(str(CurrentDirectory) + '\\Embosser_Archive\\PlateLog.txt', 'a+')
    PlateLog.write("""\n""" + """Timestamp: {:%Y-%m-%d %H:%M:%S}""".format(datetime.datetime.now()) 
    + """\n    Print Succeeded on """ + PlateType.rstrip() + """ plate. With """
    + SerialNumber.rstrip() + """ Serial Number""")
    PlateLog.close()
    return 0
  except:
    print("Couldn't write to the Platelog, but everything else went ok.")
    return

# This function defines the SQL database that we should contact in order to push and pull our information.
def SQL_cnxn():
  try:
    #print('Trying login') #debug
    server = '*******'
    database = '*******'
    username = '*******'
    password = '*******'
    driver = '{ODBC Driver 13 for SQL Server}'
    cnxn = pyodbc.connect('DRIVER=' + driver + ';SERVER=' + server + ';PORT=1443;DATABASE=' + database + ';UID=' + username + ';PWD='+ password)
    cursor = cnxn.cursor()
    return (cursor, cnxn)

  except:
    PyLogging = open(str(CurrentDirectory)+'\\Failed Plates', 'a+')
    PyLogging.write('\n' + 'Timestamp: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()) + "\n    pyodbc.connect module failed to connect to USCEVAPP006.")
    PyLogging.close()
    print ("Python couldn't establish a connection to the databse. Tell IT to look at the error log at: C:\TEST DATA\Failed\PyLogging.txt")
  
# This is the main control block. The try blocks help compensate for the failure of a PyODBC fetchone()
if __name__ == "__main__":
  print("Welcome to the Embossing program!")
  while True:
    cursor, cnxn = SQL_cnxn()
    # If possible, we try to do this with a single scan of the PSONum but this particular system needs SN
    PSONum = str.rstrip(input("Please enter PSO: ")).upper()
    SerialNumber = str.rstrip(input("Please scan the serial number: ")).upper()
    # This try block is an attempt to get away with only asking for the PSONum
    try:
      PSO_Check = cursor.execute("""Select [BOM], [Description], [TypeofPlate] from dbo.BOM where [PSONum] = '"""+PSONum+"""'""")
      BOM, Description, PlateType = cursor.fetchone()
      if PSO_Check:
        # Certain plates can have different fields, and we control for that here.
        if PlateType.upper() != 'Bb' and PlateType.upper() != 'Cc' :
          CustomerPartNumber = 'NotMuch'
          PlateMaker(PlateType, BOM, Description, SerialNumber, CustomerPartNumber)
          A_OK = PyPusher(SerialNumber, BOM, Description, PSONum)
          # If everything worked well, tell the user.
          if type(A_OK) == int:
            cnxn.close()
            print ("Seems to me like it went well.")
          # In case of failure, notify the user. Error should already be logged.
          else:
            cnxn.close()
            print ("Something didn't go quite right. Lets try that again.")


        elif str.rstrip(PlateType.upper()) == 'Bb':
          CustomerPartNumber = cursor.execute("Select [Bb NEW SAP PN] from dbo.CustPartNum where [Product_Name] = '"+BOM+"'").fetchval()
          PlateMaker(PlateType, BOM, Description, SerialNumber, CustomerPartNumber)
          A_OK = PyPusher(SerialNumber, BOM, Description, PSONum)

          if type(A_OK) == int:
            cnxn.close()
            print ("Seems to me like it went well.")
          else:
            cnxn.close()
            print ("Something didn't go quite right. Lets try that again.")

        elif str.rstrip(PlateType.upper()) == 'Cc':
          CustomerPartNumber = cursor.execute("Select [Cc Part #] from dbo.Cc where [Part #] = '"+BOM+"'").fetchval()
          PlateMaker(PlateType, BOM, Description, SerialNumber, CustomerPartNumber)
          A_OK = PyPusher(SerialNumber, BOM, Description, PSONum)

          if type(A_OK) == int:
            cnxn.close()
            print ("Seems to me like it went well.")
          else:
            cnxn.close() 
            print ("Something didn't go quite right. Lets try that again.")

        else:
          cnxn.close()
          print ("I can't find that plate type, so I can't print that plate for you.")
          time.sleep(3)

    # If none of that worked, we can try checking the secondary table.
    except:
      try:
        print("I couldn't find that PSO in the usual place.. expanding search.")
        BOM_Check = cursor.execute("Select [BOM], [Description] from dbo.BOM1 where [PSONum] = '"+PSONum+"'")
        BOM, Description = cursor.fetchone()
        PlateType = cursor.execute("Select [TypeofPlate] from dbo.BOM where [BOM] = '"+BOM+"'").fetchval()
        if BOM and Description and PlateType:
          PlateMaker(PlateType, BOM, Description, SerialNumber)
          A_OK = PyPusher(SerialNumber, BOM, Description, PSONum)
      except:
        print("I can't seem to find that PSO anywhere. I'm restarting..")
        time.sleep(1)
        PyLogging = open(str(CurrentDirectory) +'\\Embosser_Failed\\PyLogging.txt', 'a+')
        PyLogging.write('\n' + 'Timestamp: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()) + "\n    Failed, no record in BOM or BOM1. Try manual printing, and check with the ticket writer for record accuracy.")
        PyLogging.close()

