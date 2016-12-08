import win32ui
import win32print, pywintypes, win32con
from PIL import Image, ImageWin
import sys, getopt

def main(argv):
  inputfile = 'C:/Users/sci/Downloads/cat.jpg'
  try:
    opts, args = getopt.getopt(argv,"hi:",["ifile="])
  except getopt.GetoptError:
    print 'print.py -i <inputfile>'
    sys.exit(2)
  for opt, arg in opts:
    if opt == '-h':
       print 'print.py -i <inputfile>'
       sys.exit()
    elif opt in ("-i", "--ifile"):
       inputfile = arg
  print 'Input file is ', inputfile


     
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

  printer_name = win32print.GetDefaultPrinter ()
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
  devmode.PaperLength = 381
  devmode.PaperWidth = 381
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
  #printable_area = (900, 900)
  #printer_size = (900, 900)
  print "printer area", printable_area
  print "printer size", printer_size
  print "printer margins", printer_margins

  #
  # Open the image, rotate it if it's wider than
  #  it is high, and work out how much to multiply
  #  each pixel by to get it as big as possible on
  #  the page without distorting.
  #
  bmp = Image.open (file_name)
  if bmp.size[0] > bmp.size[1]:
    bmp = bmp.rotate (90)
  print "bmp size", bmp.size
  ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
  scale = min (ratios)

  scaled_width, scaled_height = [int (scale * i) for i in bmp.size]
  x1 = int ((printer_size[0] - scaled_width) / 2)
  y1 = int ((printer_size[1] - scaled_height) / 2)
  x2 = x1 + scaled_width
  y2 = y1 + scaled_height
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
   
