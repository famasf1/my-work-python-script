from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import os
import win32print
import sched
import time
import win32api

event_sche = sched.scheduler(time.time, time.sleep)

gauth = GoogleAuth()
gauth.LocalWebserverAuth()

drive = GoogleDrive(gauth)

def print_pdf(file):
    GHOSTSCRIPT_PATH = r"C:\Program Files (x86)\GhostTrap\bin\gswin32.exe"
    GSPRINT_PATH = r"C:\Program Files (x86)\GhostTrap\bin\gsprint.exe"

# YOU CAN PUT HERE THE NAME OF YOUR SPECIFIC PRINTER INSTEAD OF DEFAULT
    currentprinter = win32print.GetDefaultPrinter()

    win32api.ShellExecute(0, 'open', GSPRINT_PATH, '-ghostscript "'+ GHOSTSCRIPT_PATH +'" -dDEVICEWIDTHPOINTS=312 -dDEVICEHEIGHTPOINTS=452 -dORIENT1=true -color -dPDFFitPage -printer "'+ currentprinter + f'" "{file}"', '.', 0)
    #args = [
    #    "-dPrinted", "-dBATCH", "-dNOSAFER", "-dNOPAUSE", "-dNOPROMPT", "-dFitPage "
    #    "-q",
    #    "-dNumCopies=1",
    #    "-sDEVICE=mswinpr2",
    #    f'-sOutputFile="%printer%{win32print.GetDefaultPrinter()}"',
    #    f'"{file}"'
    #]
    #encoding = locale.getpreferredencoding()
    #args = [a.encode(encoding) for a in args]
    #ghostscript.Ghostscript(*args)


def download_Files():
    folder_id = '1EvU95WxgsYVlZqqtAdfK-W3su_vVCLSe'
    file_list = drive.ListFile({'q': "'{}' in parents and trashed=false".format(folder_id)}).GetList()
    for file1 in file_list:
        for i, file1 in enumerate(sorted(file_list, key = lambda x: x['title']), start=1):
            print('Downloading {} from GDrive ({}/{})'.format(file1['title'], i, len(file_list)))
            try:
                os.mkdir('D:\Trade in Output')
                newfolder = 'D:\Trade in Output'
            except FileExistsError:
                newfolder = 'D:\Trade in Output'
                pass
            except OSError:
                try:
                    os.mkdir('C:\Trade In Output')
                    newfolder = 'C:\Trade In Output'
                except FileExistsError:
                    newfolder = 'C:\Trade in Output'
                    pass
            except Exception:
                raise Exception
            finally:
                os.chdir(newfolder)
                file1.GetContentFile(file1['title'])
                f = os.path.join(os.getcwd(), f"{file1['title']}").replace('\\', '\\\\')
                print_pdf(f)
                event_sche.enter(15,1,download_Files,(sc,))

download_Files()
event_sche.enter(15,1,download_Files,(s,))
event_sche.run()






