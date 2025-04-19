import zipfile
import os
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar, BooleanVar
from tkinter import ttk
import xml.etree.ElementTree as ET
from glob import glob
import sys
import ctypes
import configparser

try:
    from win32com.client import Dispatch
except ImportError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Module manquant",
        "Le module 'pywin32' est requis pour créer le raccourci.\n\n"
        "Ouvre une console et tape :\n\npip install pywin32\n\n"
        "Puis :\npython -m pywin32_postinstall -install"
    )
    raise SystemExit

def creer_raccourci_si_voulu():
    appdata = os.getenv("APPDATA")
    config_dir = os.path.join(appdata, "NettoyeurGamma")
    os.makedirs(config_dir, exist_ok=True)
    config_path = os.path.join(config_dir, "config.ini")

    raccourci_dir = os.path.join(appdata, r"Microsoft\Windows\Start Menu\Programs\Scripts")
    os.makedirs(raccourci_dir, exist_ok=True)
    raccourci_path = os.path.join(raccourci_dir, "Nettoyeur Gamma.lnk")

    config = configparser.ConfigParser()
    if os.path.exists(config_path):
        config.read(config_path)
        if config.getboolean("prefs", "ne_plus_demander", fallback=False):
            return

    if os.path.exists(raccourci_path):
        return

    def demander_creation():
        fen = tk.Tk()
        fen.title("Créer un raccourci")
        fen.geometry("350x120")
        fen.resizable(False, False)

        var_ne_plus = BooleanVar(value=False)

        ttk.Label(fen, text="Voulez-vous créer un raccourci dans le menu Démarrer ?").pack(pady=10)
        ttk.Checkbutton(fen, text="Ne plus afficher ce message", variable=var_ne_plus).pack()

        btns = ttk.Frame(fen)
        btns.pack(pady=10)
        ttk.Button(btns, text="Oui", command=lambda: (fen.quit(), fen.destroy(), setattr(sys, "_create_shortcut", True), setattr(sys, "_never_ask_again", var_ne_plus.get()))).pack(side="left", padx=10)
        ttk.Button(btns, text="Non", command=lambda: (fen.quit(), fen.destroy(), setattr(sys, "_create_shortcut", False), setattr(sys, "_never_ask_again", var_ne_plus.get()))).pack(side="left", padx=10)

        fen.mainloop()

    demander_creation()

    if getattr(sys, "_never_ask_again", False):
        config["prefs"] = {"ne_plus_demander": "true"}
        with open(config_path, "w") as f:
            config.write(f)

    if getattr(sys, "_create_shortcut", False):
        shell = Dispatch('WScript.Shell')
        raccourci = shell.CreateShortCut(raccourci_path)
        if getattr(sys, 'frozen', False):
            raccourci.Targetpath = sys.executable
            raccourci.WorkingDirectory = os.path.dirname(sys.executable)
        else:
            raccourci.Targetpath = os.path.abspath(__file__)
            raccourci.WorkingDirectory = os.path.dirname(os.path.abspath(__file__))
        raccourci.IconLocation = raccourci.Targetpath
        raccourci.save()

def supprimer_blocs_pic_preencoded(layout_dir):
    supprimés = 0
    ns = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    for fichier in glob(os.path.join(layout_dir, "*.xml")):
        tree = ET.parse(fichier)
        root = tree.getroot()
        modifié = False
        for pic in root.findall(".//p:pic", ns):
            cNvPr = pic.find(".//p:nvPicPr/p:cNvPr", ns)
            if cNvPr is not None and cNvPr.get("descr") == "preencoded.png":
                parent = next((elem for elem in root.iter() if pic in list(elem)), None)
                if parent is not None:
                    parent.remove(pic)
                    supprimés += 1
                    modifié = True
        if modifié:
            tree.write(fichier, encoding='utf-8', xml_declaration=True)
    return supprimés

def traiter_pptx(fichier_entree, fichier_sortie, progress_callback=None):
    base_temp = os.path.splitext(os.path.basename(fichier_entree))[0]
    temp_dir = os.path.join(os.path.dirname(fichier_entree), f"{base_temp}_temp_process")
    try:
        if progress_callback: progress_callback(1, "Décompression en cours...")
        with zipfile.ZipFile(fichier_entree, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        layout_dir = os.path.join(temp_dir, 'ppt', 'slideLayouts')
        bloc_count = 0
        if os.path.exists(layout_dir):
            if progress_callback: progress_callback(2, "Suppression de blocs XML...")
            bloc_count = supprimer_blocs_pic_preencoded(layout_dir)

        if progress_callback: progress_callback(3, "Recompression du PPTX...")
        with zipfile.ZipFile(fichier_sortie, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    full = os.path.join(root, file)
                    rel = os.path.relpath(full, temp_dir)
                    zip_out.write(full, rel)

        if progress_callback: progress_callback(4, "Nettoyage des fichiers temporaires...")
        shutil.rmtree(temp_dir)

        return bloc_count

    except Exception as e:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        raise

class NettoyeurApp(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=20)
        self.master = master
        master.title("Nettoyeur de filigramme Gamma (.pptx)")
        master.geometry("600x220")
        master.resizable(False, False)

        self.fichier_entree = StringVar()
        self.fichier_sortie = StringVar()

        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TLabel", padding=4)

        ttk.Label(self, text="Fichier PPTX d'entrée :").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self, textvariable=self.fichier_entree, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(self, text="Parcourir...", command=self.choisir_fichier).grid(row=0, column=2)

        ttk.Label(self, text="Fichier PPTX de sortie :").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(self, textvariable=self.fichier_sortie, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(self, text="Enregistrer sous...", command=self.choisir_sortie).grid(row=1, column=2)

        self.progress = ttk.Progressbar(self, orient='horizontal', length=500, mode='determinate', maximum=4)
        self.progress.grid(row=2, column=0, columnspan=3, pady=15)
        self.status_label = ttk.Label(self, text="Prêt.")
        self.status_label.grid(row=3, column=0, columnspan=3, sticky=tk.W)

        self.run_btn = ttk.Button(self, text="Lancer le nettoyage", command=self.start_processing)
        self.run_btn.grid(row=4, column=0, columnspan=3, pady=0)

        self.grid()

    def choisir_fichier(self):
        f = filedialog.askopenfilename(filetypes=[("Fichier PPTX", "*.pptx")])
        if f:
            self.fichier_entree.set(f)
            default_out = f.replace(".pptx", "_modifie.pptx")
            self.fichier_sortie.set(default_out)

    def choisir_sortie(self):
        f = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("Fichier PPTX", "*.pptx")])
        if f:
            self.fichier_sortie.set(f)

    def update_progress(self, step, message):
        self.progress['value'] = step
        self.status_label.config(text=message)
        self.master.update_idletasks()

    def start_processing(self):
        if not self.fichier_entree.get() or not self.fichier_sortie.get():
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier d'entrée et un fichier de sortie.")
            return
        self.run_btn.state(['disabled'])
        threading.Thread(target=self.run_task, daemon=True).start()

    def run_task(self):
        try:
            bloc_count = traiter_pptx(
                self.fichier_entree.get(),
                self.fichier_sortie.get(),
                self.update_progress
            )
            self.show_bilan(bloc_count)
        except Exception as e:
            messagebox.showerror("Erreur", str(e))
            self.status_label.config(text="Erreur lors du traitement.")
        finally:
            self.run_btn.state(['!disabled'])

    def show_bilan(self, bloc_count):
        result = messagebox.showinfo(
            "Terminé", f"{bloc_count} bloc(s) supprimé(s) avec succès.")
        self.progress['value'] = 0
        self.status_label.config(text="Prêt.")

if __name__ == "__main__":
    creer_raccourci_si_voulu()
    root = tk.Tk()
    app = NettoyeurApp(root)
    root.mainloop()