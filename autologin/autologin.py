# -*- coding: utf-8 -*-
'''
Created on 1 Septembre 2012: 'autologin.py' 
-Utilisation du module PyWinAuto pour piloter chrome est aller sur des sites bancaires\
-Utilisation cv2 = OpenCV pour comparer des images\
@author: Python4d
'''
from time import sleep
from pywinauto.application import Application #http://pywinauto.googlecode.com/hg/pywinauto/docs/index.html
from pywinauto.win32functions import SetCursorPos
import cv2

class banques:
    """Classe principale permettant de regrouper les fonctions d'automatisation"""
    def __init__(self):
        "Constructeur de l'objet principale du module pywinauto: l'applicaton a automatiser"
        self.app = Application()
        #Etape 0 (cf blog www.python4d.com)
        try: #Connection de l'objet app avec le process lié à la fenêtre principale Google Chrome
            self.app.connect_(title_re=".* - Google Chrome", class_name=r"Chrome_WidgetWin_1") 
        except: #lancer un nouveau process Chrome et s'y connecter
        #Etape 1(cf blog www.python4d.com)
            self.app.start_(r"C:\Users\damien\AppData\Local\Google\Chrome\Application\chrome.exe", timeout=2)
            sleep(5)
        self.app["Chrome_WidgetWin_1"].Maximize()
        #Mettre le curseur souris à un endroit qui ne perturbe pas le focus des fenêtres du browser
        SetCursorPos(0, 0)
        sleep(1)
        #fenêtre de l'application tout entière Chrome
        self.ChromeWin = self.app["Chrome_WidgetWin_1"]
        #fenêtre (textbox) de l'URL
        self.ChromeUrl = self.app["Chrome_WidgetWin_1"]["Chrome_OmniboxView"]
        #Etape n°2 (cf blog www.python4d.com)
        self.newtab()

    def newtab(self):
        "Méthode qui lance un CTRL-T sur l'application Chrome: création d'un nouvel onglet"
        sleep(2)
        self.ChromeWin.TypeKeys(r"^T") 
        #Mettre le curseur souris à un endroit qui ne perturbe pas le focus des fenêtres du browser
        SetCursorPos(0, 0)
        sleep(1)
        
    def matchtemplate(self, img1, img2):
        "Méthode de comparaison entre deux images [cv2.matchTemplate (OpenCV)]"
        imgVignette=cv2.imread(img2)
        imgEcran=cv2.imread(img1)
        result=cv2.matchTemplate(imgEcran,imgVignette,5)
        (_, _, _, maxLoc) = cv2.minMaxLoc(result, mask=None)
        return maxLoc
        
    def ingdirect(self, compte="12345{TAB}01011901{ENTER}", mdp="012345"):
        "Methode automatisant les entrées claviers et souris pour le site INGDIRECT"
        dirbase = r".//ingdirect//"
        adrbanque = r"https://secure.ingdirect.fr/public/displayLogin.jsf"
        #Etape n°3 (cf blog www.python4d.com)
        self.ChromeUrl.TypeKeys("^a" + adrbanque + "{ENTER}"),sleep(4)
        #Etape n°4 (cf blog www.python4d.com)
        self.ChromeWin.TypeKeys(compte),sleep(5)
        #Etape n°5 (cf blog www.python4d.com)
        WinBoursoramaImage = self.ChromeWin.CaptureAsImage()
        WinBoursoramaImage.save(dirbase + "WebImage.png")
        coordy = [0,]*10
        coordx = [0,]*10
        x, y = self.matchtemplate(dirbase + "WebImage.png", dirbase + "VALIDER.png")
        #Etape n°6/9 (cf blog www.python4d.com)
        for i in range(0, 10):
            small = str(i) + '.png'
            coordx[i], coordy[i] = self.matchtemplate(dirbase + "WebImage.png", dirbase + small)
        premiere_absisse = x #absisse du premier code correspond à l'absisse de la forme "VALIDER". 
        #Etape n°7/9 (cf blog www.python4d.com)
        for i in range(3):
          codex, codey = self.matchtemplate(dirbase + "WebImage.png", dirbase + "code.png")
          code=int(round(abs(premiere_absisse-codex)/32.0)) #on considère que chaque emplacement de code est séparé par 32 points (! dépend de la résolution !)
          self.ChromeWin.ClickInput(coords=(coordx[int(mdp[code])] + 10, coordy[int(mdp[code])] + 10), double=False)
          WinBoursoramaImage = self.ChromeWin.CaptureAsImage()
          WinBoursoramaImage.save(dirbase + "WebImage.png")
        #Etape n°8 (cf blog www.python4d.com)
        self.ChromeWin.ClickInput(coords=(x + 10, y + 10), double=False)  

if __name__ == "__main__":
  #connection sur l'application chrome et création d'un nouvel ongletAvertissement
  banque=banques()
  banque.ingdirect("12345{TAB}02021968{ENTER}","987654")
  