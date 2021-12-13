import multiprocessing
import os
import win32com.client
import easygui
import time
from multiprocessing import set_start_method, Pool

def unlockExcel(filename1, filename2, password):
    print(filename1 + " is being worked on now!")
    xcl = win32com.client.Dispatch("Excel.Application")
    wb = xcl.Workbooks.Open(filename1, False, False, None, password, password, True)
    xcl.DisplayAlerts = False
    wb.SaveAs(filename2, None, '', '')
    print(filename1 + " is unlocked now!")
    #xcl.Quit()

def closeFile():

    try:
        os.system('TASKKILL /F /IM excel.exe')

    except Exception:
        print("KU")

if __name__=="__main__":

    set_start_method("spawn")

    easygui.msgbox(
        "Program is initiating. \nAll existing Excel processes will be shut down.\nPlease take the opportunity to save "
        "and close existing Excel files.\nYou can stop the program by exiting the windows console.\n")

    pass_word = easygui.passwordbox("enter your password to unlock", title='Unlocking Models')

    #starttime = time.time()

    i = 0
    t = '_unlocked'

    pool = Pool(multiprocessing.cpu_count())

    for entry in os.scandir(os.getcwd()):

        if (entry.path.endswith(".xlsx") or entry.path.endswith(".xlsb")) and entry.is_file():
            try:
                filename1 = entry.path
                filename2 = entry.path[:-5] + t + entry.path[-5:]
                pool.apply_async(unlockExcel, (filename1, filename2, pass_word))
                i += 1
            except:
                easygui.msgbox(
                    "Password is incorrect or unknown error is encountered. Try again, before contacting Ray.",
                    title='Error Alert')
                pool.close()
                pool.join()
                closeFile()
                #endtime = time.time()
                #duration = endtime - starttime
                #print(str(duration))
                os.sys.exit(1)

    pool.close()
    pool.join()

    closeFile()

    #endtime = time.time()

    #duration = endtime - starttime

    #print(str(duration))

    easygui.msgbox(
        "Job is completed with " + str(
            i) + " file(s) unlocked.\n" + "Feedbacks are welcomed to Ray at starstream521@gmail.com",
        title='Job Finished!')

    os.sys.exit(0)
