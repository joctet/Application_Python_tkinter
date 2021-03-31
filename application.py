import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import tkinter.font as font
import pandas as pd
import os
def popupmsg(msg):
    popup = tk.Tk()
    popup.wm_title("Votre attention s'il vous plait !")
    popup.configure(background="#FA7F7F")
    label_popup = tk.Label(popup, text=msg, foreground = "#641E16", background = "#FA7F7F")
    label_popup['font'] = f_label
    label_popup.pack()
    bouton_popup = tk.Button(popup, text="C'est compris!", command = popup.destroy)
    bouton_popup.configure(foreground="#641E16", background = "white")
    bouton_popup['font'] = f_bouton
    bouton_popup.pack()
    popup.mainloop()
def traitement():
    chemin_fichier = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xlsx')), ('CSV', '*.csv')])
    if chemin_fichier.endswith('.csv'):
        df = pd.read_csv(chemin_fichier)
    else:
        df = pd.read_excel(chemin_fichier, engine='openpyxl')
    chemin_dossier = os.path.dirname(chemin_fichier)
    if "données_modifiées.xlsx" in os.listdir(chemin_dossier):
        popupmsg("Il y a déjà un fichier nommé 'données_modifiées.xlsx' dans le dossier, veuillez renommer ce fichier pour que le programme puisse en générer un nouveau.")
    else :
        df.to_excel(str(chemin_dossier) + "\\" + "données_modifiées.xlsx", index=False)
        popupmsg("Votre fichier excel est généré =) Vous pouvez le trouver dans le même dossier que le fichier excel initial." + "\n" + "\n" + "A bientôt!" + "\n")

interface = tk.Tk()
interface.title("Programme de modification de fichier Excel")
interface.configure(background="#2B00FA")
# 1er message
f_label = font.Font(family='Times New Roman', size=20)
f_bouton = font.Font(family='Times New Roman', size=15, weight="bold")
label = tk.Label(text="\n"  + "Bonjour! Vous avez lancé le programme d'aide à la modification de fichier Excel." + "\n"  + "\n" + "Veuillez sélectionner un fichier Excel ou CSV :" + "\n" , foreground = "white", background = "#2B00FA")
label['font'] = f_label
label.pack()
# 1er bouton
bouton = tk.Button(text='Cliquez ici pour charger un fichier', command=traitement)
bouton.place(relx=0.200, rely=0.06, height=500, width=147)
bouton.configure(foreground="#2B00FA")
bouton['font'] = f_bouton
bouton.pack(expand="yes")
# 2eme message
label2 = tk.Label(text="\n"  + "Une fois le fichier séléctionné, le programme va générer un nouveau fichier excel (nommé 'Données_modifiées.xlsx') dans le même dossier que le fichier séléctionné." + "\n"  + "\n" + "Pour relancer le programme, veuillez renommer le fichier 'Données_modifiées.xlsx' pour ne pas rentrer en conflit avec le programme." + "\n", foreground = "white", background = "#2B00FA")
label2['font'] = f_label
label2.pack()
# 2eme bouton
bouton2 = tk.Button(interface, text="J'ai fini merci!", command = interface.destroy)
bouton2.configure(foreground="#2B00FA")
bouton2['font'] = f_bouton
bouton2.pack()
interface.mainloop()