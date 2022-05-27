from __future__ import print_function
import ctypes
from ctypes import wintypes
import psutil
import win32gui
import win32con
import win32api
#from win32.win32gui import FindWindow, GetWindowRect, MoveWindow
#from win32 import win32gui
#import win32ui, win32con, win32api
from collections import namedtuple
import sys,os,logging
import time
from datetime import datetime

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#If there is no such folder, the script will create one automatically
folder_location_log = r'log'
if not os.path.exists(folder_location_log):os.mkdir(folder_location_log)

logging.basicConfig(filename=datetime.now().strftime('log\Ominis-Monitor-Application%d-%m-%Y-%H-%M-%S.log'), filemode='w', format='%(asctime)s - %(message)s', level=logging.INFO,datefmt='%Y-%m-%d %H:%M:%S')
logging.info('The Script for Monitoring Omnini Apllication is Started')

user32 = ctypes.WinDLL('user32', use_last_error=True)


if not hasattr(wintypes, 'LPDWORD'): # PY2
    wintypes.LPDWORD = ctypes.POINTER(wintypes.DWORD)

WindowInfo = namedtuple('WindowInfo', 'pid title')
WNDENUMPROC = ctypes.WINFUNCTYPE(
    wintypes.BOOL,
    wintypes.HWND,    # _In_ hWnd
    wintypes.LPARAM,) # _In_ lParam


user32.EnumWindows.argtypes = (
   WNDENUMPROC,      # _In_ lpEnumFunc
   wintypes.LPARAM,) # _In_ lParam

user32.IsWindowVisible.argtypes = (
    wintypes.HWND,) # _In_ hWnd

user32.GetWindowThreadProcessId.restype = wintypes.DWORD
user32.GetWindowThreadProcessId.argtypes = (
  wintypes.HWND,     # _In_      hWnd
  wintypes.LPDWORD,) # _Out_opt_ lpdwProcessId

user32.GetWindowTextLengthW.argtypes = (
   wintypes.HWND,) # _In_ hWnd

user32.GetWindowTextW.argtypes = (
    wintypes.HWND,   # _In_  hWnd
    wintypes.LPWSTR, # _Out_ lpString
    ctypes.c_int,)   # _In_  nMaxCount


psapi = ctypes.WinDLL('psapi', use_last_error=True)

psapi.EnumProcesses.argtypes = (
   wintypes.LPDWORD,  # _Out_ pProcessIds
   wintypes.DWORD,    # _In_  cb
   wintypes.LPDWORD,) # _Out_ pBytesReturned

def list_windows():
    '''Return a sorted list of visible windows.'''
    logging.info('list of visible windows inside function started')
    result = []
    @WNDENUMPROC
    def enum_proc(hWnd, lParam):
        if user32.IsWindowVisible(hWnd):
            pid = wintypes.DWORD()
            tid = user32.GetWindowThreadProcessId(
                        hWnd, ctypes.byref(pid))
            length = user32.GetWindowTextLengthW(hWnd) + 1
            title = ctypes.create_unicode_buffer(length)
            user32.GetWindowTextW(hWnd, title, length)
            result.append(WindowInfo(pid.value, title.value))
        return True
    user32.EnumWindows(enum_proc, 0)
    logging.info('list of visible windows inside function ended')
    return sorted(result)
   
def monitor_app(app_exe):
    logging.info('Checking if '+serviceName+' is running or not')
    for process in psutil.process_iter():
        if process.name() == app_exe:
            return True
def send_mail():
    # me == my email address
    # you == recipient's email address
    me = "my email address"
    you = "recipient's email address"
    
    text = "Hi!\nHow are you?\nHere is the link you wanted:\nhttp://www.python.org"
    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Issue with Omini Application:Might be Hanged"
    msg['From'] = me
    msg['To'] = you


    html = """\
    <html>
      <head></head>
      <body>
        <p>Hello Team<br>
           There is issue with your application <br>
           <br><br>
           Regards,<br>
           Omini Monitoring Team
        </p>
      </body>
    </html>
    """

    # Record the MIME types of both parts - text/plain and text/html.
    part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')

    # Attach parts into message container.
    # According to RFC 2046, the last part of a multipart message, in this case
    # the HTML message, is best and preferred.
    msg.attach(part1)
    msg.attach(part2)
    # Send the message via local SMTP server.
    mail = smtplib.SMTP('yoursmtpserver', 587)

    mail.ehlo()

    mail.starttls()

    mail.login('youremail', 'yourpassword')
    logging.info('Triggering mail part is started: '+you)
    mail.sendmail(me, you, msg.as_string())
    mail.quit()

if __name__ == '__main__':
    logging.info('list of visible windows function calling started')
    serviceName='HDSentinel.exe'
    consoleName='Hard Disk Sentinel'
    #consoleNameDll='Hard Disk Sentinel Evaluation Version'
    consoleNameDll='Advanced Power Management'
    serviceNameStatus=monitor_app(serviceName)
    listOfVisibleWindows=list_windows()
    #print('listOfVisibleWindows',listOfVisibleWindows)
    omniCosnsoleVisibleStatus=omniCosnsoleDLLVisibleStatus=''
    for a in listOfVisibleWindows:
        print('title',a.title)
        if a.title==consoleName:
            omniCosnsoleVisibleStatus='Yes'
        if a.title==consoleNameDll:
            omniCosnsoleDLLVisibleStatus='Yes'
            handleConsole = win32gui.FindWindow(None, a.title)
            hangedConsoleDLLTitle=a.title
            hangedConsoleDLLPID=a.pid
            
    if ((serviceNameStatus) and (omniCosnsoleVisibleStatus=='Yes') and (omniCosnsoleDLLVisibleStatus=='')):
        logging.info(serviceName+' is  up and running')
        print(serviceName,'is up and running')
    elif ((serviceNameStatus) and (omniCosnsoleVisibleStatus=='Yes') and (omniCosnsoleDLLVisibleStatus=='Yes')):
        logging.info(serviceName+' is  hanged and going to be closed ')
        #send_mail()        
        try:
                win32gui.PostMessage(handleConsole, win32con.WM_CLOSE, 0, 0)
                logging.info('the Window is closed: '+hangedConsoleDLLTitle)
                time.sleep(10)
        except Exception as e:
            print("exception occured",e)
            logging.info('Exception occured, forcefully termination '+str(e))
            handleConsoleDLLForce = win32api.OpenProcess(win32con.PROCESS_TERMINATE, 0, hangedConsoleDLLPID)
            if handleConsoleDLLForce:
                win32api.TerminateProcess(handleConsoleDLLForce,0)
                win32api.CloseHandle(handleConsoleDLLForce)
                logging.info('the Window is closed forcefully: '+hangedConsoleDLLTitle)
            #pass
    elif not serviceNameStatus:
        logging.info(serviceName+' is not running')
        print(serviceName,'is not running and triggerring email')
        #send_mail()
        #logging.info('The Script for Monitoring Omnini Apllication is ended')
    else:
        logging.info('no condition is not met')
        
logging.info('The Script for Monitoring Omnini Apllication is ended')        
        
                