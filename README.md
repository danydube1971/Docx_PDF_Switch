# Docx_PDF_Switch

Script Python permettant de convertir des fichiers DOCX en PDF et des fichiers PDF en DOCX en lots à l'aide d'une interface graphique simple.

![Docx_PDF_Switch](https://github.com/danydube1971/Docx_PDF_Switch/assets/74633244/7a5f58ad-4095-4e47-b542-cb74d1d217cb)

*Testé dans Linux Mint 21.3*

# Prérequis

**Python 3.11 (ou une version compatible) :** Assurez-vous que Python est installé sur votre système. 

**LibreOffice :** Utilisé pour la conversion DOCX en PDF. 

**Modules Python :** tkinter, subprocess, pdf2docx, pypandoc. 

# Installation

Étape 1 : Installation de LibreOffice LibreOffice est nécessaire pour convertir des fichiers DOCX en PDF. 
LibreOffice est nécessaire pour convertir des fichiers DOCX en PDF. 

**Linux (Debian/Ubuntu)**

`sudo apt-get install libreoffice`

**MacOS (via Homebrew)**

`brew install libreoffice`

Étape 2 : Installation des Modules Python

Installez les modules Python nécessaires en utilisant pip.

`pip install pdf2docx pypandoc`

Étape 3 : Téléchargement et Configuration du Script

Téléchargez le script : Sauvegardez le code ci-dessous dans un fichier nommé *DocxPDFSwitch.py*

Étape 4 : Utilisation du Script

Exécution du Script : Ouvrez un terminal (ou une invite de commande) et exécutez le script avec la commande suivante :

`python DocxPDFSwitch.py`
    
   Interface Graphique : Une fenêtre s'ouvre avec deux boutons principaux :
       
  DOCX en PDF : Cliquez sur ce bouton pour sélectionner un dossier contenant des fichiers DOCX à convertir en PDF.
       
  PDF en DOCX : Cliquez sur ce bouton pour sélectionner un dossier contenant des fichiers PDF à convertir en DOCX.
  
  Suivi de la Progression : Une fenêtre de progression s'affiche, montrant l'état de la conversion des fichiers. La fenêtre se ferme automatiquement une fois la conversion terminée.
  
  Quitter l'Application : Cliquez sur le bouton "Quitter" pour fermer l'application.

  

    
       
    
