'''
  ****************************************************************************************************************
  *                                         Zebra label maker program                                            *
  ****************************************************************************************************************
  * This program takes PSONum as an argument and uses it in a query to a server's BOM table.                     *
  * It uses PySQLCaller() to pull the information from there and displays it to the user (on the off chance that *
  * they can troubleshoot for themselves). It then plugs some of this query's values into PyTicketWriter().      *
  * From there it pulls up the master image and writes the current information on it. It then sends the Quantity *
  * from the pull to the zebra printer.                                                                          *
  *                                                                                                              *
  * This program should use roughly 10 GB per year, and will store the successes in the "Label Archive" folder.  *
  *                                                                                                              *
  * It is very important not to alter the master image, as the program uses it to write labels.                  *
  *                                                                                                              *
  * This program needs to run from its own folder, so try using shortcuts for users.                             *
  *                                                                                                              *
  * The calls here are to a specific database, it was written to track industrial blowers but they can be        *
  * rewritten to pretty much whatever you want.                                                                  *
  *                                                                                                              *
  * I left the debug statements in, just commented. I hope they're helpful to illustrate possible fail points.   *
  *                                                                                                              *
  ****************************************************************************************************************
  '''

import textwrap
import code128
import pyodbc
import os
import datetime
import time
import random
import win32api
import win32print
#These PIL explicit imports are necessary.
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw

#If the files for logging success & failure don't exist, make them.
if not os.path.exists('Label Archive'):
    os.mkdir("Label Archive")
if not os.path.exists('Failed Labels'):
    os.mkdir("Failed Labels")

##These are the global constants used by the functions below
#This defaults the labelpath variable to your current folder
labelpath = os.path.dirname(os.path.realpath(__file__))
#I separated the master images from the rest of the folder. This is the path to those.
master_path = str(labelpath) + "\\Master Images\\"
#If label making goes well, push stuff here.
archive_path = str(labelpath) + "\\Label Archive"
#If we can't print a label, or can print but with errors, put those here.
fail_path = str(labelpath) + "\\Failed Labels"
#This program was written to print to this specific printer, though the win32api will support
#printing to just about any printer as long as you have the right name.
#If you need to print to a different label maker, I suggest
#changing the size values in PyTicketWriter() to suit your needs.
printer_name = 'ZDesigner ZT230-300dpi ZPL'
#These are used for global variables later
BC_Holder = ''
TypeofPlate = ''

#This function queries the db in the SQL_cnxn function, and passes the relevent information to other functions
def PySQLCaller(PSONum):
  #This uses pyodbc to return a curson and cnxn so that we can play with SQL (don't forget to fill in your SQL values!)
  cursor, cnxn = SQL_cnxn()

  #check for pre-existing records on dbo.Blowers (see if its been through the system already)
  try:
    #Run this check. If they exist on the table, it runs a challenge. Otherwise it will pass control
    #directly to the 'if PSO_Lookup()' statement.
    CheckVSBlowers = cursor.execute("Select [PSO], [SN] from dbo.Blowers where [PSO] = '" + PSONum + "' and SN is not null")
    #We only want to know if these combinations exist, so any amount of them is enough. fetchone() is faster.
    SNCheck = cursor.fetchone()
    #If we can't find it in the SQL db, great. Keep going.
    if SNCheck is None:
      pass
    #We found existing records of this. Confirmation dialog that we want to redo this record.
    else:
      #This loop is just to trap the user in to giving valid input.
      while True:
        duplicate_challenge = input ("Part or all of that PSO has already been through the system. Do you want to reprint them? [Y/N] " )
        if duplicate_challenge.upper() in ['Y', 'N']:
            break
            #move on to yes or no criteria.
        else:
            #Trap user into inputting valid options.
            print('That is not a valid option, please type "Y" or "N".')
      #Continue to print stuff, user knows it's a reprint.
      if duplicate_challenge.upper() == 'Y':
        pass
      #Drop out of PySQLCaller and back into main while loop. The user doesn't want to reprint these tickets.
      elif duplicate_challenge.upper() == 'N':
        return

  #This rarely triggers, it would have to be querying a table without PSO and SN or some similar problem.
  except:
    PyLogging = open(fail_path, 'a+')
    PyLogging.write('\n' + 'Timestamp: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()) + "\n    Verification failed. Try looking at ODBC driver revisions or connections to USCEVAPP006 and make sure the relevant fields exist.")
    PyLogging.close()
    #Even if this error triggers, we want to try to print this PSO.
    print ("Python couldn't verify the originality of the PSO you entered. Continuing anyway.")   
  
  #Try looking up the PSONum, if its already been put in
  if PSO_Lookup(PSONum):
    ##TODO: inelegant, ugly, wasteful.
    #This query will almost definitely not reflect your own. Rewrite it to reflect what you want to query.
    BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, Quantity, Notes = PSO_Lookup(PSONum)
    #Tell the user what we found, so if there are any errors, they have a better idea of why it failed.
    print ("These are the fields I'm printing:")
    print (PSONum, ",".lstrip(" "), BOM, ",", Customer, ",", Description, ",", Model, ",", Size, ",", RP_Spec, ",", Shaft, ",", Discharge, ",", Quantity, ",", Notes) #debug

  #Try looking it up by BOM. This fails if some fields aren't entered correctly.
  elif BOM_Lookup(BOM_Check()):
    #BOM is the UID of this particular SQL db so if it exists, it just doesn't have the right PSO associated with it.
    #That's super easy to fix, we'll just use the BOM record we have and then slap the right PSO and quantity on it when we're done.
    try:
      BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, Notes = BOM_Lookup(globals()['BC_Holder'])
      Quantity = int(input ("I found that BOM in the database. What Quantity should we run this in? "))
      cursor.execute("""UPDATE dbo.BOM SET [QUANTITY] = '"""+str(Quantity)+"""', [PSONUM] = '"""+PSONum+"""' where [BOM] like '"""+BOM+"""'""")
      cnxn.commit()
      print ("These are the fields I'm printing:")
      print (PSONum, ",".lstrip(" "), BOM, ",", Customer, ",", Description, ",", Model, ",", Size, ",", RP_Spec, ",", Shaft, ",", Discharge, ",", Quantity, ",", Notes) #debug
      BOM = globals()['BC_Holder']
    except:
      pass

  #If we can't find the record by PSO or BOM, switch to manual input
  else:
    #This doesn't assign to anything, it quits out when its done on its own.
    PyManualInput(PSONum, BOM_Check())

  #We should have our values now, one way or another.
  try:
    #I tried this a bunch of different ways, and this seemed the most consistent.
    #The BOM should be a given. If not, you have an error message.
    if BOM:
      #print ("BOM1 insert with BOM") #debug
      #This calls the ticket writing program to use the information we have.
      PyTicketWriter(PSONum, BOM, Customer, Description, Size, RP_Spec, Shaft, Discharge, Quantity, Notes)
      #After we print the ticket(s), we then check the secondary table. This was done for a series of reasons the client came up with.
      BOM1 = cursor.execute("""SELECT [BOM] from BOM1 where PSONum = '"""+PSONum+"""'""").fetchval()
      if BOM1:
        #print("Found BOM1") # debug
        pass
      else:
        #print("Trying BOM1 insert") # debug
        #Secondary BOM table insert if we can't find preexisting records. Again, done at the request of the client.
        cursor.execute("""INSERT INTO BOM1 (
          BOM, Description, TypeofPlate, Model, Size, RP_Spec, Shaft, Discharge, Customer, PSONum, Quantity, [TimeStamp], Notes)
          SELECT BOM, Description, TypeofPlate, Model, Size, RP_Spec, Shaft, Discharge, Customer, PSONum, Quantity, GETDATE(), Notes 
          FROM BOM where PSONum = '"""+PSONum+"""'""")
        cnxn.commit()
        cnxn.close()
      return (Quantity)

    #Handling for situations that we didn't anticipate. Kept a global variable for this specific purpose. Wasn't used much, but works fine.
    else:
      #print("BOM1 insert without BOM") #debug
      #Assign BOM from our global if it exists.
      BOM = globals()['BC_Holder']
      #We need to print out labels in this case as well.
      PyTicketWriter(PSONum, BOM, Customer, Description, Size, RP_Spec, Shaft, Discharge, Quantity, Notes, Notes)
      #Now that we've printed the label, we run similar checks for the BOM1 table.
      BOM1 = cursor.execute("""SELECT [BOM] from BOM1 where PSONum = '"""+PSONum+"""'""").fetchval()
      #Found a secondary record? Cool. Do nothing.
      if BOM1:
        pass
      #Can't find a record on the secondary BOM table? Lets push in the information we have.
      else:
        cursor.execute("""INSERT INTO BOM1 SELECT BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, PSONum, Quantity, Notes FROM BOM where PSONum ='"""+PSONum+"""'""")
        cnxn.commit()
      cnxn.close()
      return (Quantity)

  #Couldn't find a BOM, couldn't compensate. Soft fail.
  except:
    #print("BOM1 insert failed.") # debug
    cnxn.close()
    return ()

# This function takes the information that we pass to it and overlways it on top of the master image. In this
# case it is writing those on top of a label for a production management system. It uses an inelegant system
# to narrow down which picture it should use based on the orientation of the shaft & discharge. It then goes
# into a loop that prints copies of the ticket based on the quantity presented. The client in this case wanted
# to randomly mark 10% of certain blowers for testing, and there's a block for that. Then it tries printing
# to the named printer in the global variables block. Failing that, it pushes to the default printer.
def PyTicketWriter(PSONum, BOM, Customer, Description, Size, RP_Spec, Shaft, Discharge, Quantity, Notes):
  #print ("trying Ticket") #debug
  try:
    #set base number for label iterator, we start at 1 so we can print at least 1 label.
    number_of_labels = 1
    #open template image to write on
    img = Image.open(str(master_path) + "master_template.jpg")
    draw = ImageDraw.Draw(img)
    #I initially used a bunch of font sizes, turned out unnecessary but I left them in for good measure.
    #The default font of the Zebra printer is a true type font called swissel. The others are variants.
    ##font = ImageFont.truetype(<font-file>, <font-size>)
    font60 = ImageFont.truetype("swissb.ttf", 60)
    font45 = ImageFont.truetype("swissel.ttf", 45)
    font35 = ImageFont.truetype("swissel.ttf", 35)
    font28 = ImageFont.truetype("swissel.ttf", 28)
    font14 = ImageFont.truetype("swissck.ttf", 14)  
    font7 = ImageFont.truetype("swissck.ttf", 7)
    #These statements overlay the image for each variable passed to the function
    ##draw.text((x, y),"Text Here",(r,g,b))
    draw.text((11, 213),PSONum,(0,0,0),font=font45)
    draw.text((11, 28),BOM,(0,0,0),font=font60)
    draw.text((11, 304),Customer,(0,0,0),font=font35)
    draw.text((11, 115),textwrap.fill(Description, width=34),(0,0,0),font=font28)
    draw.text((11, 465),Notes,(0,0,0),font=font45)
    draw.text((352, 300),RP_Spec,(0,0,0),font=font45)
    #print ("I see Shaft = "+Shaft+"\nand Discharge = "+Discharge+"") # debug
    ##TODO: replace with dict
    #This functional piece of code selects an image based on the variables we found in the SQL tables.
    if Shaft == "LEFT" and Discharge == "*":
      Lpic = Image.open(str(master_path) + "Lpic.jpg")
      img.paste(Lpic, (458, 500))
      Lpic.close()
    elif Shaft == "RIGHT" and Discharge == "*":
      Rpic = Image.open(str(master_path) + "Rpic.jpg")
      img.paste(Rpic, (458, 500))
      Rpic.close()
    elif Shaft == "TOP" and Discharge == "*":
      Tpic = Image.open(str(master_path) + "Tpic.jpg")
      img.paste(Tpic, (458, 500))
      Tpic.close()      
    elif Shaft == "BOTTOM" and Discharge == "*":
      Bpic = Image.open(str(master_path) + "Bpic.jpg")
      img.paste(Bpic, (458, 500))
      Bpic.close()
    elif Shaft == "LEFT" and Discharge == "TOP":
      LTpic = Image.open(str(master_path) + "LTpic.jpg")
      img.paste(LTpic, (458, 500))
      LTpic.close()
    elif Shaft == "LEFT" and Discharge == "BOTTOM":
      LBpic = Image.open(str(master_path) + "LBpic.jpg")
      img.paste(LBpic, (458, 500))
      LBpic.close()
    elif Shaft == "RIGHT" and Discharge == "TOP":
      RTpic = Image.open(str(master_path) + "RTpic.jpg")
      img.paste(RTpic, (458, 500))
      RTpic.close()
    elif Shaft == "RIGHT" and Discharge == "BOTTOM":
      RBpic = Image.open(str(master_path) + "RBpic.jpg")
      img.paste(RBpic, (458, 500))
      RBpic.close()
    elif Shaft == "TOP" and Discharge == "RIGHT":
      TRpic = Image.open(str(master_path) + "TRpic.jpg")
      img.paste(TRpic, (458, 500))
      TRpic.close()
    elif Shaft == "TOP" and Discharge == "LEFT":
      TLpic = Image.open(str(master_path) + "TLpic.jpg")
      img.paste(TLpic, (458, 500))
      TLpic.close()
    elif Shaft == "BOTTOM" and Discharge == "RIGHT":
      BRpic = Image.open(str(master_path) + "BRpic.jpg")  
      img.paste(BRpic, (458, 500))
      BRpic.close()
    elif Shaft == "BOTTOM" and Discharge == "LEFT":
      BLpic = Image.open(str(master_path) + "BLpic.jpg")
      img.paste(BLpic, (458, 500))
      BLpic.close()
    else:
      print ("Couldn't find a Shaft & Discharge picture.. printing anyway.")
    #iterator for printing correct quantity of labels
    while number_of_labels <= Quantity:
      #print ("Trying iterator") #debug 
      #We're going to name our temporary image with the same convention as the one we'll see when we name the pdf.
      #PSO123456-2_of_4.pcx
      new_pcx_path = os.path.join(labelpath, str(PSONum) + "-" + str(number_of_labels) + "_of_" + str(Quantity) + '.pcx')
      #We're not going to mess with the master, so we want to pull up a copy to put edits on.
      unique_img = img.copy()
      unique_draw = ImageDraw.Draw(unique_img)
      #The client wanted to randomize certain blower labels with an "X" that marked it for testing. This does that.
      MTR_randomizing = random.randrange(1,10)
      #Barcoding module for BOM value. We're using code128 because of its versatility and fairly good reads. Also, the printer
      #sometimes gets misaligned so we put it excessively away from the edge of the label to compensate.
      BOM_Barcode = code128.image(str(PSONum))
      unique_img.paste(BOM_Barcode, (115, 366))
      #This randomizes testing for the blowers sized 36, 33, and 32 to a 1:10 chance.
      #print ("Starting Randomizing.") #debug
      #If we have blowers of these sizes AND we get a 1:10 chance, then we mark it with an "X"
      if Size == ('36' or '33' or '32') and MTR_randomizing == 1:
          #print ("Randomization Passed.") #debug
          unique_draw.text((99, 520),"X",(0,0,0),font=font60)
      #print ("Show statement.") #debug
      #Tried for the highest quality resize here. This was a point of struggle but
      #in the end this worked okay.
      unique_resize = unique_img.resize((609, 609), Image.LANCZOS)
      #unique_img.show() #debug
      #unique_resize.show() #debug
      #Name the image using our convention and save image as PDF in the label_archive folder.
      new_pdf_path = os.path.join(archive_path, str(PSONum) + "-" + str(number_of_labels) + "_of_" + str(Quantity) + '.pdf')
      unique_img.save(new_pdf_path)
      #Print statement for printer_name printer. ZT300 MUST be set to 3in x 3in or it will look awful.
      try:
        win32api.ShellExecute (0,"print",new_pdf_path,'/d:"%s"' % win32print.OpenPrinter(printer_name),".",0)
      except:
        print("Couldn't find the "+printer_name+" so I'm printing to the default printer.")
        win32api.ShellExecute (0,"print",new_pdf_path,'/d:"%s"' % win32print.GetDefaultPrinter (),".",0)
      #Push our iterator up one, and then retry the loop.
      number_of_labels = number_of_labels + 1
      #Close our temp image. This is so we try the randomization on different images.
      unique_img.close()

  except:
    #TODO needs logger for .pcx opening and draw failure
    print("I couldn't write to the printer.\nAre you connected to the "+printer_name+" printer, or have you selected a default printer that I can print to?")

  #No matter what happens, we don't want to keep our images open in memory.
  img.close()

# This function queries the SQL database for this specific record by the provided PSONum. In this case if the
# record can't be found it queries the secondary database. Again, this was client specified functionality. It
# should probably be taken out of the try block for single-query functionality.
def PSO_Lookup(PSONum):
  # Grab the SQL connection
  cursor, cnxn = SQL_cnxn()
  try:
    # Query primary database and grab first record, then stuff it into PSO_Check
    PSO_Query = cursor.execute("Select [BOM], [Customer], [Description], [Model], [Size], [RP_Spec], [Shaft], [Discharge], [Quantity], [Notes] from dbo.BOM where [PSONum] = '"+PSONum+"'")
    PSO_Check = cursor.fetchone()
    BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, Quantity, Notes = PSO_Check
    # Purely for redundancy
    globals()['BC_Holder'] = BOM
    # This is to catch any fields that happen to be null. This particular problem tends to screw a lot of
    # the latter parts of this program up, so it defaults to manual input.
    if PSO_Check:
      for field in PSO_Check:
        if field == None:
          # If a field is Null, switch to manual
          # Manual does input for both BOM and BOM1
          BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, TypeOfPlate, Quantity, PSONum, Notes = PyManualInput(PSONum, BOM)
          return 1
        else:
          #print ("Parsed a field.") # debug
          continue
      #print ("Passed PSO_Check.") # debug
      BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, Quantity, Notes = PSO_Check
      return (BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, Quantity, Notes)
    else:
      print ("Couldn't find a valid PSO.")
      pass
  except:
    # Same as above, just for a different database. PyODBC fails if there's no record whatsoever so it was encased
    # in a try block.
    PSO_Query = cursor.execute("Select [BOM], [Customer], [Description], [Model], [Size], [RP_Spec], [Shaft], [Discharge], [Quantity], [Notes] from dbo.BOM1 where [PSONum] = '"+PSONum+"'")
    PSO_Check = cursor.fetchone()
    BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, Quantity, Notes = PSO_Check
    globals()['BC_Holder'] = BOM
    if PSO_Check:
      for field in PSO_Check:
        if field == None:
          #If a field is Null, switch to manual
          #Manual does input for both BOM and BOM1
          BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, TypeOfPlate, Quantity, PSONum, Notes = PyManualInput(PSONum, BOM)
          return 1
        else:
          #print ("Parsed a field.") # debug
          continue
      #print ("Passed PSO_Check.") # debug
      BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, Quantity, Notes = PSO_Check
      return (BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, Quantity, Notes)
    else:
      print ("Couldn't find a valid PSO.")
      pass

# Simple check for BOM for manual input.
def BOM_Check():
  BOM_input = str.rstrip(input("I couldn't find that PSO in the Database.\nCan you please enter the BOM (part number) below letter for letter?\n: ")).upper()
  globals()['BC_Holder'] = BOM_input
  return (BOM_input)

# Part of another redundancy. If there's a record of the BOM but not a record of the PSONum, this will look
# it up and return the values to be combined with the PSONum so it can be pushed to the ticket.
def BOM_Lookup(BOM):
  cursor, cnxn = SQL_cnxn()
  try:
    # Find the first (hopefully only) record associated with this BOM and stuff it into BOM_Return. Similar functionality to PSO_Lookup.
    BOM_Query = cursor.execute("Select [BOM], [Customer], [Description], [Model], [Size], [RP_Spec], [Shaft], [Discharge], [Notes] from dbo.BOM where [BOM] like '"+BOM+"'")
    BOM_Return = cursor.fetchone()
    BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, Notes = BOM_Return
    if BOM_Return:
      for field in BOM_Return:
        # Make sure that none of the fields are null, and if they are, switch to manual.
        if field == None:
          BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, TypeOfPlate, Quantity, PSONum, Notes = PyManualInput(PSONum, BOM)
          # Now draw out the ticket and print it
          PyTicketWriter(PSONum, BOM, Customer, Description, Size, RP_Spec, Shaft, Discharge, Quantity, Notes)
          return 1
        else:
          continue
    BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, Notes = BOM_Return
    globals()['BC_Holder'] = BOM
    return (BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, Notes)
  except:
    pass

# This function holds the information to access the SQL database. You'll want to fill this in with something
# other than asterisks, unless you don't actually want to access any databases.
def SQL_cnxn():
  try:
    server = '*******'
    database = '*******'
    username = '*******'
    password = '*******'
    driver = '{ODBC Driver 13 for SQL Server}'
    #table = 'dbo.SlipTable'
    cnxn = pyodbc.connect('DRIVER=' + driver + ';SERVER=' + server + ';PORT=1443;DATABASE=' + database + ';UID=' + username + ';PWD='+ password)
    cursor = cnxn.cursor()
    return (cursor, cnxn)
  except:
    PyLogging = open(fail_path, 'a+')
    PyLogging.write('\n' + 'Timestamp: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()) + "\n    pyodbc.connect module failed to connect to USCEVAPP006.")
    PyLogging.close()
    print ("Python couldn't establish a connection to the databse. Tell IT to look at the error log at: C:\TEST DATA\Failed\PyLogging.txt")

# The final failsafe for getting a label printed. If there is no record at all, this provides prompts to create
# a new one manually. It then inserts it into both the primary and secondary databases.
def PyManualInput(PSONum, BOM):
  cursor, cnxn = SQL_cnxn()
  # Tell the user whats going on.
  print ('''I couldn't find a valid record of this blower, but I did find a BOM: "'''+BOM+'''".\nSwitching to manual input.''')
  # This is a series of blocks, one for each field. Each one tries to fetch a value from the database in case we got here because
  # PSO_Lookup or BOM_Lookup found a null field in the SQL database and queries the user as to the correctness of this field.
  # If there's no record, it defaults to a manual input with a trapping while loop for confirmation.
  while True:
    # This one is for the 'description' field.
    try:
      Description = cursor.execute("Select [Description] from dbo.BOM where [BOM] like '"+globals()['BC_Holder']+"'").fetchval().upper()
      print("I found the description: "+Description+"")
      Desc_Challenge = input('Is this correct? (Type "Y" or "N") ')
      while True:
        if Desc_Challenge.upper() in ['Y', 'N']:
            break
            # Move on to yes or no criteria.
        else:
            # Trap user into inputting valid options.
            print('That is not a valid option, please type "Y" or "N".')
      # Continue to print
      if Desc_Challenge.upper() == 'Y':
        break
      # Drop out of PySQLCaller and back into main while
      elif Desc_Challenge.upper() == 'N':
        Description = input("What is the Description for this PSO? ").upper()
        break
    except:
      Description = input("What is the Description for this PSO? ").upper()
      break

  # This one is for the 'Customer' field.
  while True:
    try:
      Customer =  cursor.execute("Select [Customer] from dbo.BOM where [BOM] like '"+globals()['BC_Holder']+"'").fetchval().upper()
      print("I found the Customer: "+Customer+"")
      Cust_Challenge = input('Is this correct? (Type "Y" or "N") ')
      while True:
        if Cust_Challenge.upper() in ['Y', 'N']:
            break
            # Move on to yes or no criteria.
        else:
            # Trap user into inputting valid options.
            print('That is not a valid option, please type "Y" or "N".')
      # Continue to print
      if Cust_Challenge.upper() == 'Y':
        break
      # Drop out of PySQLCaller and back into main while
      elif Cust_Challenge.upper() == 'N':
        Customer = input("What is the Customer for this PSO? ").upper()
        break
    except:
      Customer = input("What is the Customer for this PSO? ").upper()
      break

  # This one is for the 'Model' field.
  while True:
    try:
      Model = cursor.execute("Select [Model] from dbo.BOM where [BOM] like '"+globals()['BC_Holder']+"'").fetchval().upper()
      print("I found the Model: "+Model+"")
      while True:
        Model_Challenge = input('Is this correct? (Type "Y" or "N") ')
        if Model_Challenge.upper() in ['Y', 'N']:
            break
            # Move on to yes or no criteria.
        else:
            # Trap user into inputting valid options.
            print('That is not a valid option, please type "Y" or "N".')
      # Continue to print
      if Model_Challenge.upper() == 'Y':
        break
      # Drop out of PySQLCaller and back into main while
      elif Model_Challenge.upper() == 'N':
        Model = input("What is the Model for this PSO? ").upper()
        break
    except:
      Model = input("What is the Model for this PSO? ").upper()
      break

  # This one is for the 'Size' field.
  while True:
    try:
      Size = cursor.execute("Select [Size] from dbo.BOM where [BOM] like '"+globals()['BC_Holder']+"'").fetchval()
      print("I found the Size: "+Size+"")
      while True:
        Size_Challenge = input('Is this correct? (Type "Y" or "N") ')
        if Size_Challenge.upper() in ['Y', 'N']:
            break
            # Move on to yes or no criteria.
        else:
            # Trap user into inputting valid options.
            print('That is not a valid option, please type "Y" or "N".')
      # Continue to print
      if Size_Challenge.upper() == 'Y':
        break
      # Drop out of PySQLCaller and back into main while
      elif Size_Challenge.upper() == 'N':
        Size = input("What is the Size for this PSO? ")
        break
    except:
      Size = input("What is the Size for this PSO? ")
      break

  # This one is for the 'RP_Spec' field.
  while True:
    try:
      RP_Spec = cursor.execute("Select [RP_Spec] from dbo.BOM where [BOM] like '"+globals()['BC_Holder']+"'").fetchval().upper()
      print("I found the RP_Spec: "+RP_Spec+"")
      RP_Spec_Challenge = input('Is this correct? (Type "Y" or "N") ')
      while True:
        if RP_Spec_Challenge.upper() in ['Y', 'N']:
            break
            # Move on to yes or no criteria.
        else:
            # Trap user into inputting valid options.
            print('That is not a valid option, please type "Y" or "N".')
      # Continue to print
      if RP_Spec_Challenge.upper() == 'Y':
        break
      # Drop out of PySQLCaller and back into main while
      elif RP_Spec_Challenge.upper() == 'N':
        RP_Spec = input("What is the RP_Spec for this PSO? ").upper()
        break
    except:
      RP_Spec = input("What is the RP_Spec for this PSO? ").upper()
      break

  # This one is for the 'Shaft' field.
  while True:
    try:
      Shaft = cursor.execute("Select [Shaft] from dbo.BOM where [BOM] like '"+globals()['BC_Holder']+"'").fetchval().upper()
      print("I found the Shaft: "+Shaft+"")
      Shaft_Challenge = input('Is this correct? (Type "Y" or "N") ')
      while True:
        if Shaft_Challenge.upper() in ['Y', 'N']:
            break
            # Move on to yes or no criteria.
        else:
            # Trap user into inputting valid options.
            print('That is not a valid option, please type "Y" or "N".')
      # Continue to print
      if Shaft_Challenge.upper() == 'Y':
        break
      # Drop out of PySQLCaller and back into main while
      elif Shaft_Challenge.upper() == 'N':
        Shaft = input("What is the Shaft for this PSO? ").upper()
        break
    except:
      Shaft = input("What is the Shaft for this PSO? ").upper()
      break

  # Discharge manual entry block
  while True:
    try:
      Discharge = cursor.execute("Select [Discharge] from dbo.BOM where [BOM] like '"+globals()['BC_Holder']+"'").fetchval().upper()
      print("I found the Discharge: "+Discharge+"")
      Discharge_Challenge = input('Is this correct? (Type "Y" or "N") ')
      while True:
        if Discharge_Challenge.upper() in ['Y', 'N']:
            break
            # Move on to yes or no criteria.
        else:
            # Trap user into inputting valid options.
            print('That is not a valid option, please type "Y" or "N".')
      # Continue to print
      if Discharge_Challenge.upper() == 'Y':
        break
      # Drop out of PySQLCaller and back into main while
      elif Discharge_Challenge.upper() == 'N':
        Discharge = input("What is the Discharge for this PSO? ").upper()
        break
    except:
      Discharge = input("What is the Discharge for this PSO? ").upper()
      break

  # Notes manual entry block
  while True
    try:
      Notes = cursor.execute("Select [Notes] from dbo.BOM where [BOM] like '"+globals()['BC_Holder']+"'").fetchval().upper()
      print("I found the Notes: "+Notes+"")
      Notes_Challenge = input('Is this correct? (Type "Y" or "N") ')
      while True:
        if Notes_Challenge.upper() in ['Y', 'N']:
            break
            # Move on to yes or no criteria.
        else:
            # Trap user into inputting valid options.
            print('That is not a valid option, please type "Y" or "N".')
      # Continue to print
      if Notes_Challenge.upper() == 'Y':
        break
      # Drop out of PySQLCaller and back into main while
      elif Notes_Challenge.upper() == 'N':
        Notes = input("What is the Discharge for this PSO? ").upper()
        break
    except:
      Notes = input("What is the Discharge for this PSO? ").upper()
      break
  if len(str(Notes)) < 1:
    Notes = 'None'

  # Type of plate manual entry block
  while True:
    TypeofPlate = input("What type of plate should this unit use? \nChoices are Aa, Bb, Cc, Dd, Ee: ").upper()
    globals()['TypeOfPlate'] = TypeofPlate
    if TypeOfPlate == 'AA' or TypeOfPlate == 'BB' or TypeOfPlate == 'CC' or TypeOfPlate == 'DD' or TypeOfPlate == 'EE':
      break
    else:
      print("I didn't get a type of plate I recognize. Can you reenter the plate type please?")
  BOM = globals()['BC_Holder']

  Quantity = int(input("How many of these should I print? "))
  #print ('BC_Holder is: '+str(globals()['BC_Holder'])) #debug
  # Two flavors of input statement. One is for each database, in case the first input fails.
  try:
    #print("trying manual BOM insert") #debug
    Manual_Insert = cursor.execute("""INSERT INTO dbo.BOM ([BOM], [Customer], [Description], [Model], [Size], [RP_Spec], [Shaft], [Discharge], [TypeofPlate], [Quantity], [PSONum], [Notes])
      VALUES ('"""+str(globals()['BC_Holder'])+"""', '"""+Customer+"""', '"""+Description+"""', '"""+Model+"""', '"""+str(Size)+"""', '"""+RP_Spec+"""', '"""+Shaft+"""', '"""+Discharge+"""', '"""+TypeofPlate+"""', '"""+str(Quantity)+"""', '"""+PSONum+"""', '"""+Notes+"""')""")
    cnxn.commit()
    cnxn.close()
    #print("Successful manual insert for BOM.") #debug
  except:
    #print("trying secondary manual BOM insert") #debug
    Manual_update = cursor.execute("""UPDATE dbo.BOM SET
    [Customer] = '"""+Customer+"""', [Description] = '"""+Description+"""', [Model] = '"""+Model+"""', 
    [Size] = '"""+str(Size)+"""', [RP_Spec] = '"""+RP_Spec+"""', [Shaft] = '"""+Shaft+"""', [Discharge] = '"""+Discharge+"""', 
    [TypeofPlate] = '"""+TypeofPlate+"""', [Quantity] = '"""+str(Quantity)+"""', 
    [PSONum] = '"""+PSONum+"""', [Notes] = '"""+Notes+"""' WHERE BOM = '"""+globals()['BC_Holder']+"""'""")
    cnxn.commit()
    cnxn.close()
    #print("Successful secondary manual insert for BOM.") #debug
  return (BOM, Customer, Description, Model, Size, RP_Spec, Shaft, Discharge, TypeOfPlate, Quantity, PSONum, Notes)

#TODO: control flow here is unclear, rework.
if __name__ == '__main__':
  print ("Welcome to the label printing program!")
  ##TODO log amend for logins
  #This while loop contains the entirety of the program and is here to make the program continue to run
  while True:
    #This is the input box for the user, stripped to prevent mismatches with the sql database.
    #Users saw their input as unclear so I stripped the sides.
    PSONum = str.rstrip(input("Please enter the PSO number you'd like to print: ")).upper()
    #If the user put something in, yay, continue doing stuff.
    if PSONum:
      #This is actually the end of this run of the program. Control flows to PySQLCaller. If the
      #program was successful, no quantity is given and an error message pops out.
      Qty = PySQLCaller(PSONum)
      #If successful, Qty should give an int value.
      if type(Qty) == int:
        print ("You just sent", Qty, "labels of", PSONum, "to the printer.")
      else:
        #TODO descriptive message & logging
        print ("Something didn't go quite right. Lets try that again.")
    
    #If the user doesn't enter anything on the PSONum prompt and hits return, the program quits.
    else:
      print ("Thanks for using the label printing program, have a nice day!")
      time.sleep(3)
      quit()
