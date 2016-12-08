win32printer

A tool to print via the commandline on windows.

# Install 

[Install python 2.7](https://www.python.org/downloads/release/python-2712/)

Add environment variables on Windows for C:\Python27, C:\Python27\Scripts


Install libs

```
pip install pypiwin32
pip install Pillow
```

# Run

```
C:\Python27\python.exe C:\path\to\win32print\print.py -i "C:\path\to\picture.jpg"
```


# Run via linux

This tool was intended to use a printer that does not have drivers for linux.

We installed a VirtualBox on linux, and then run print commands with winexe.

## Install
Install winexe

```
sudo apt-get install winexe
```

Authorize remote desktop on windows.

## Run
And then you can run 

```
winexe -U HOME/admin%password //192.168.1.xx 'C:\Python27\python.exe C:\path\to\win32print\print.py -i "C:\path\to\picture.jpg"'
```

## troubleshooting
If you see authorization error, try 

[turn off UAC remote restrictions](https://support.microsoft.com/en-us/kb/951016)

if it still does not work, try

[turn off UAC completely](http://www.howtogeek.com/howto/4820/how-to-really-completely-disable-uac-on-windows-7/)

if it still does not work, try to disable firewalls, and make the network private instead of public.
