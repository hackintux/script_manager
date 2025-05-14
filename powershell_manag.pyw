import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import ctypes
import urllib.request

# === CONFIGURATION COULEURS ===
FOND_GENERAL = "#d0f0c0"
FOND_BOUTON = "#66bb6a"
FOND_BOUTON_HOVER = "#43a047"
FOND_TEXTE = "#ffffff"
FOND_RESULTAT = "#ffffff"
COULEUR_TEXTE = "#2e7d32"

# === SCRIPT GITHUB ===
SCRIPT_GITHUB_URL = "https://raw.githubusercontent.com/hackintux/Gestionnaire-IT/main/script.ps1"

# === EXE GITHUB INSTALL ===
EXE_URL = "https://raw.githubusercontent.com/hackintux/caps_lock/main/caps_lock.exe"
EXE_NOM = "caps_lock.exe"
EXE_DEST = os.path.join(os.environ["ProgramFiles"], "CapsLock", EXE_NOM)
DOSSIER_DEMARRAGE = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs", "Startup")

# === FONCTIONS ===
def est_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def relancer_en_admin():
    if not est_admin():
        script_path = os.path.abspath(sys.argv[0])
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, f'"{script_path}"', None, 1
        )
        sys.exit()

def afficher_resultats(titre, texte):
    fenetre = tk.Toplevel(root, bg=FOND_GENERAL)
    fenetre.title(titre)
    fenetre.geometry("700x500")
    zone = scrolledtext.ScrolledText(fenetre, font=("Segoe UI", 10), bg=FOND_RESULTAT, fg=COULEUR_TEXTE, insertbackground=COULEUR_TEXTE)
    zone.pack(fill="both", expand=True, padx=20, pady=20)
    zone.insert(tk.END, texte)
    zone.config(state="disabled")

def executer_script_depuis_github():
    try:
        temp_path = os.path.join(os.getenv("TEMP"), "cliconline_git.ps1")
        urllib.request.urlretrieve(SCRIPT_GITHUB_URL, temp_path)
        subprocess.run(
            ["powershell.exe", "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass", "-File", script],
            capture_output=True,
            text=True
        )
        sortie = result.stdout + "\n" + result.stderr
        afficher_resultats("Gestionnaire IT ClicOnLine (GitHub)", sortie)
        os.remove(temp_path)
    except Exception as e:
        messagebox.showerror("Erreur", f"Échec du téléchargement ou de l'exécution : {e}")

def charger_et_executer():
    path = filedialog.askopenfilename(defaultextension=".ps1", filetypes=[("PowerShell Scripts", "*.ps1")])
    if path:
        try:
            result = subprocess.run(["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", path], capture_output=True, text=True)
            sortie = result.stdout + "\n" + result.stderr
            afficher_resultats(f"Sortie de {os.path.basename(path)}", sortie)
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

def installer_et_lancer_exe():
    try:
        dossier_install = os.path.dirname(EXE_DEST)
        os.makedirs(dossier_install, exist_ok=True)
        urllib.request.urlretrieve(EXE_URL, EXE_DEST)

        shortcut_path = os.path.join(DOSSIER_DEMARRAGE, "CapsLock.lnk")
        import pythoncom
        from win32com.client import Dispatch
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = EXE_DEST
        shortcut.WorkingDirectory = dossier_install
        shortcut.IconLocation = EXE_DEST
        shortcut.save()

        subprocess.Popen(EXE_DEST, shell=True)
        messagebox.showinfo("Succès", "Installation et lancement terminés.")

    except Exception as e:
        messagebox.showerror("Erreur", f"Échec de l'installation : {e}")

# === FORCÉMENT REDEMARRÉ EN ADMIN AU LANCEMENT ===
relancer_en_admin()

# === INTERFACE ===
root = tk.Tk()
root.title("ClicOnLine")
root.geometry("450x350")
root.configure(bg=FOND_GENERAL)

style = ttk.Style()
style.theme_use("clam")
style.configure("TButton",
                font=("Segoe UI", 11, "bold"),
                padding=12,
                foreground=FOND_TEXTE,
                background=FOND_BOUTON,
                borderwidth=0)
style.map("TButton",
          background=[("active", FOND_BOUTON_HOVER)],
          foreground=[("active", FOND_TEXTE)])

frame = tk.Frame(root, bg=FOND_GENERAL)
frame.pack(expand=True, fill="both", padx=20, pady=20)

titre = tk.Label(frame, text="Gestionnaire IT - ClicOnLine", font=("Segoe UI", 18, "bold"), bg=FOND_GENERAL, fg=COULEUR_TEXTE)
titre.pack(pady=10)

ttk.Button(frame, text="▶ Démarrer Gestionnaire IT", command=executer_script_depuis_github).pack(fill="x", pady=6)
ttk.Button(frame, text="📂 Charger un programme", command=charger_et_executer).pack(fill="x", pady=6)
ttk.Button(frame, text="⬇ Installer & lancer Caps Lock", command=installer_et_lancer_exe).pack(fill="x", pady=6)

root.mainloop()