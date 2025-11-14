import os
import sys
import time
import gc
from habetools import Bot
from datetime import datetime
import importlib

class Menuitem:
    def __init__(self, command, label, nbparams, ret):
        self.command = command        
        self.label = label
        self.nbparams = nbparams
        self.ret = ret

rootApp = os.getcwd()
hardgreen = "\033[32m\033[1m"
normalgreen = "\033[32m\033[2m"
normalcolor = "\033[0m"

def matrix_print(text, speed=0.01):
    for char in text:
        sys.stdout.write(char)
        sys.stdout.flush()
        time.sleep(speed)
    print()

def dotail(profil):
    logFilename = "{0}{1}log{1}{3}{2}.log".format(rootApp, os.path.sep, profil, dnow)
    os.system("tail -f {0}".format(logFilename))

def mencol(nb, fonc, comment):
    return "{0}{3} - {4} {1} - {5}{2}".format(hardgreen, normalgreen, normalcolor, nb, fonc, comment)

def drkcol(str):
    return "{0}{2}{1}".format(hardgreen, normalcolor, str)

def clear():
    return os.system('clear')

nb_args = len(sys.argv)
jsonfilefromarg = "default" if (nb_args == 1) else sys.argv[1]
mode_menu = "default" if (nb_args < 3) else sys.argv[2]
clear()

bot = Bot()
bot.init_main(jsonfilefromarg)

matrix_print(drkcol("\nHi Neo, I'm the Habetool"))
matrix_print(drkcol("Your wish is my order\n"))
matrix_print(drkcol("What I can do for you :\n"))
menulist = []
menulist.append(Menuitem("test", "test", 0, False))
menulist.append(Menuitem("copyxl","copy xls (see conf file)", 0, False))
menulist.append(Menuitem("search","search", 0, True))
for idx, menuitem in enumerate(menulist):
    print (mencol(idx, menuitem.command, menuitem.label))
    if menuitem.ret:
        print(drkcol("#####"))
print(drkcol("#####"))
print(mencol("66", "advanced", "advanced menu"))
print(mencol("93", "editparams", "edit default.json"))
print(mencol("99", "exit", "exit this menu"))
dothat = input(drkcol("\n\nReady to rock : "))
today = datetime.now()
dnow = today.strftime(r"%y%m%d")
if dothat == "66":
    mode_menu = "advanced"
if dothat == "93":
    print(drkcol("edit params -r\n"))
    os.system("nano data/conf/default.json")    
if dothat == "99":
    print(drkcol("\nsee you soon, Neo\n"))
    del bot
    gc.collect
    quit()
try:
    if int(dothat) < 50:
        cmdstr = "nop"
        item = menulist[int(dothat)]
        cmd = item.command
        prms = int(item.nbparams)
        prmcmdlist = []
        for i in range(prms):
            prmcmdlist.append(input(drkcol(f"enter param {i} :")))
        prm2 = "" if (len(prmcmdlist) < 1) else prmcmdlist[0]
        prm3 = "" if (len(prmcmdlist) < 2) else prmcmdlist[1]            
        bot.main(cmd, jsonfilefromarg, prm2, prm3)            
except Exception as e:
    print (e)
    print(f"\n{hardgreen}bad command (something went wrong){normalcolor}\n")
    