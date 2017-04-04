import win32ui
import win32print, pywintypes, win32con
from PIL import Image, ImageWin
import sys, getopt
import subprocess

def help():
    print 'print.py -i <inputfile>'
    print '-i <inputfile>, --ifile <inputfile>'
    print '\t input file to print'
    print '-f <folder>, --sambafolder <folder>'
    print '\t use samba if the input file is on a shared samba folder'
    print '\t the samba path should be in the form ' + r'\\192.168.1.152\smilecooker' 
    print '-u <user>, --sambauser <user>'
    print '\t samba user should be in the form ' + r'whatEver\myuser'
    print '\t if you run into 1332 error, try adding ' + r'whatEver\ before user'
    print '-p <pass>, --sambapassword <pass>'
    print '\t samba user password'
    print '-d <letter>, --sambadriveletter <letter>'
    print '\t mounting letter for the samba folder, should be in the form "x:"'
    print '-h <printer>, --printer <printer>'
    print '\t name of the printer. Default is windows default printer.'
    print '-w <width>, --width <width>'
    print '\t Paper width in 0.1mm'
    print '-l <length>, --length <length>'
    print '\t Paper length in 0.1mm'
    print '-r <angle>, --rotate <angle>'
    print '\t Rotate media of angle in degree: 90, 180, ...'
    print '-o <value>, --object-fit <value>'
    print '\t Similar to css object-fit. Available values are \'cover\' or \'contain\''
    print '\nThis script is developed to be used from linux.'
    print 'You can install this script on windows, and run it from a linux bash with a command like: '
    print r'winexe -U HOME/windowsuser%windowspassword //192.168.1.153 \'C:\Python27\python.exe C:\Python27\Scripts\print.py -i "x:\mypicture.jpg" --sambafolder "\\192.168.1.152\smilecooker" --sambauser "whatEver\user" --sambapassword "password" --sambadriveletterd "x:"\''
   
def main(argv):
  inputfile = r'C:\Users\sci\Downloads\dbch-print.jpg'
  bUseSambaFolder = False
  sambaFolder = r'\\192.168.1.152\smilecooker'
  sambaUser = r'whatEver\guest'
  sambaPassword = r'guest'
  sambaDriveLetter = r'x:'
  printer_name = ''
  paperWidth = 360
  paperLength = 360
  rotate = 0
  objectFit = 'cover'
  
  try:
    opts, args = getopt.getopt(argv,"hi:f:u:p:d:h:w:l:r:o:",["ifile=", "sambafolder=", "sambauser=", "sambapassword=", "sambadriveletter=", "printer=", "width=", "length=", "rotate=", "object-fit="])
  except getopt.GetoptError:
    help()
    sys.exit(2)
  for opt, arg in opts:
    if opt == '-h':
       help()
       sys.exit()
    elif opt in ("-i", "--ifile"):
       inputfile = arg
    elif opt in ("-f", "--sambafolder"):
       sambaFolder = arg
       bUseSambaFolder = True
    elif opt in ("-u", "--sambauser"):
       sambaUser = arg
    elif opt in ("-p", "--sambapassword"):
       sambaPassword = arg
    elif opt in ("-d", "--sambadriveletter"):
       sambaDriveLetter = arg
    elif opt in ("-h", "--printer"):
       printer_name = arg
    elif opt in ("-w", "--width"):
       paperWidth = int(arg)
    elif opt in ("-l", "--length"):
       paperLength = int(arg)
    elif opt in ("-r", "--rotate"):
       rotate = int(arg)
    elif opt in ("-o", "--object-fit"):
       if arg == 'cover' or arg == 'contain':
           objectFit = arg
       else:
           print 'Value ' + arg + ' is not available for object-fit. Possible values are \'cover\' or \'contain\''
  print 'Input file is ', inputfile
  sambaCommand = r'net use ' + sambaDriveLetter + ' ' + sambaFolder + ' /user:' + sambaUser + ' ' + sambaPassword
  if bUseSambaFolder:
    print 'Input samba command is ', sambaCommand
    subprocess.call(sambaCommand, shell=True)
  if printer_name == '':
    printer_name = win32print.GetDefaultPrinter ()
     
  #
  # Constants for GetDeviceCaps
  #
  #
  # HORZRES / VERTRES = printable area
  #
  HORZRES = 8
  VERTRES = 10
  #
  # LOGPIXELS = dots per inch
  #
  LOGPIXELSX = 88
  LOGPIXELSY = 90
  #
  # PHYSICALWIDTH/HEIGHT = total area
  #
  PHYSICALWIDTH = 110
  PHYSICALHEIGHT = 111
  #
  # PHYSICALOFFSETX/Y = left / top margin
  #
  PHYSICALOFFSETX = 112
  PHYSICALOFFSETY = 113
  
  file_name = inputfile

  #
  # You can only write a Device-independent bitmap
  #  directly to a Windows device context; therefore
  #  we need (for ease) to use the Python Imaging
  #  Library to manipulate the image.
  #
  # Create a device context from a named printer
  #  and assess the printable size of the paper.
  #

  printer = win32print.OpenPrinter(printer_name, {'DesiredAccess': win32print.PRINTER_ALL_ACCESS})
  d = win32print.GetPrinter(printer, 2)
  devmode = d['pDevMode']
  #print 'Status ', d['Status']
  #for n in dir(devmode):
  #  print "%s\t%s" % (n, getattr(devmode, n))
  #if d[18]:
    #print "Printer not ready"
  #print ':'.join(x.encode('hex') for x in devmode.DriverData)

  devmode.PaperLength = paperWidth  # in 0.1mm
  devmode.PaperWidth = paperLength # in 0.1mm
  # paper for square sticker
  # devmode.PaperLength = 381 # in 0.1mm
  # devmode.PaperWidth = 381 # in 0.1mm
  # paper for hello my name is
  #devmode.PaperLength = 2400  # in 0.1mm
  #devmode.PaperWidth = 1800 # in 0.1mm
  # paper for flipbook
  #devmode.PaperLength = 411  # in 0.1mm
  #devmode.PaperWidth = 1079 # in 0.1mm
  win32print.SetPrinter(printer, 2, d, 0)

###  dmsize=win32print.DocumentProperties(0, printer, printer_name, None, None, 0)
### dmDriverExtra should be total size - fixed size
##  driverextra=dmsize - pywintypes.DEVMODEType().Size  ## need a better way to get DEVMODE.dmSize
##  dm=pywintypes.DEVMODEType(driverextra)
##  #win32print.DocumentProperties(0, printer, printer_name, dm, None, win32con.DM_IN_BUFFER)
##  #for n in dir(dm):
##  #  print "%s\t%s" % (n, getattr(dm, n))
##  #dm.Fields=dm.Fields|win32con.DM_ORIENTATION|win32con.DM_COPIES
##  #dm.Orientation=win32con.DMORIENT_LANDSCAPE
##  #dm.Copies=2
##  #dm.PaperSize = 256
##  #dm.PaperLength = 381
##  #dm.PaperWidth = 381
##  for n in dir(dm):
##    print "%s\t%s" % (n, getattr(dm, n))
##  win32print.DocumentProperties(0, printer, printer_name, dm, dm, win32con.DM_IN_BUFFER|win32con.DM_OUT_BUFFER)
##  for n in dir(dm):
##    print "%s\t%s" % (n, getattr(dm, n))
    
  hDC = win32ui.CreateDC ()
  hDC.CreatePrinterDC (printer_name)
  printable_area = hDC.GetDeviceCaps (HORZRES), hDC.GetDeviceCaps (VERTRES)
  printer_size = hDC.GetDeviceCaps (PHYSICALWIDTH), hDC.GetDeviceCaps (PHYSICALHEIGHT)
  printer_margins = hDC.GetDeviceCaps (PHYSICALOFFSETX), hDC.GetDeviceCaps (PHYSICALOFFSETY)
  #printable_area = (350*3, 1412*3)
  #printer_size = (350*3, 1412*3)
  print "printer area", printable_area
  print "printer size", printer_size
  print "printer margins", printer_margins

  #
  # Open the image, rotate it if it's wider than
  #  it is high, and work out how much to multiply
  #  each pixel by to get it as big as possible on
  #  the page without distorting.
  #
  open (file_name)
  bmp = Image.open (file_name)
  #if bmp.size[0] > bmp.size[1]:
  #  bmp = bmp.rotate (90)
  #  bmp = bmp.rotate (180)
  if rotate != 0:
      bmp = bmp.rotate (rotate)
  print "bmp size", bmp.size
  ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
  # fit to page
  if objectFit == "contain":
    scale = min (ratios)
  elif objectFit == "cover":
    scale = max (ratios)
  
  scaled_width, scaled_height = [int (scale * i) for i in bmp.size]
  bmp = bmp.resize((scaled_width, scaled_height), Image.HAMMING)
  x1 = int ((printer_size[0] - scaled_width) / 2)
  y1 = int ((printer_size[1] - scaled_height) / 2)
  x2 = x1 + scaled_width
  y2 = y1 + scaled_height

  # align right
  #x1 = (printer_size[0] - scaled_width)
  #x2 = x1 + scaled_width
  #scale = printer_size[1] / bmp.size[0]
  #x1 = 0
  #y1 = 0
  #x2 = 360
  #y2 = 446
  
  print "print rect: ", x1, y1, x2, y2
  #sys.exit()
  
  #
  # Start the print job, and draw the bitmap to
  #  the printer device at the scaled size.
  #
  try:
    hDC.StartDoc (file_name)
    hDC.StartPage ()

    dib = ImageWin.Dib (bmp)

    dib.draw (hDC.GetHandleOutput (), (x1, y1, x2, y2))

    hDC.EndPage ()
    hDC.EndDoc ()
    hDC.DeleteDC ()
  except win32ui.error as e:
    print "Unexpected error:", e

if __name__ == "__main__":
   main(sys.argv[1:])
   
