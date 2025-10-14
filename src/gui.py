import tkinter as tk
import os
from tkinter import ttk
from tkinter import filedialog, messagebox
from pathlib import Path, PureWindowsPath
from install_logic import Menueintrag, Verknüpfung_erstellen, finden, finden_, entfernen, install_prog, Verknüpfung_löschen, Menueintrag_entf
from install_logic import PROGRAM_NAME, ICON
import threading
import ctypes
import sys
import shutil
import time


class Installationsassistent:
    def __init__(self, callback=None):
        self.root = tk.Tk()
        self.root.title("Installationsassistent für Windows")
        self.root.geometry("450x350")

        self.zustand = "start"
        self.callback = callback

        self._create_widgets()
        self._layout_widgets()

        self.progressbar = ttk.Progressbar(self.root, orient='horizontal', mode='determinate', length=250)

        self.root.protocol("WM_DELETE_WINDOW", self.__xbeenden)

        self.label3 = tk.Label(self.root, font=("Arial", 11), pady=0, justify="left")

        self.__menue()


    def __menue(self):
        hauptmenu = tk.Menu(self.root)
        self.root.config(menu=hauptmenu)

        #dateimenu erstellen
        dateimenu = tk.Menu(hauptmenu, tearoff=0) # menu selber wird erstellt
        hauptmenu.add_cascade(label="Datei", menu=dateimenu)    #Menureiter "Datei" wird erstellt
        dateimenu.add_command(label="Beenden", command=self.__beenden)

        hilfemenu = tk.Menu(hauptmenu, tearoff=0)
        hauptmenu.add_cascade(label="Hilfe", menu=hilfemenu)
        hilfemenu.add_command(label="Über", command=self.__ueber)


    def run(self):
        self.root.mainloop()


    def _create_widgets(self):
        self.t = None
        self.var1 = tk.BooleanVar(value=True)
        self.check1 = tk.Checkbutton(self.root, text="Menüeintrag erstellen", variable=self.var1)

        self.var2 = tk.BooleanVar(value=False)
        self.check2 = tk.Checkbutton(self.root, text="Desktop Verknüpfung erstellen", variable=self.var2)

        pfad = Path(r"C:\Program Files") / PROGRAM_NAME
        pfad_win = str(PureWindowsPath(pfad))
        self.entry_var = tk.StringVar(value=pfad_win)
        self.entry = tk.Entry(self.root, textvariable=self.entry_var)

        self.open_btn = tk.Button(self.root, text="Öffnen", command=self.__öffen_dataidialog)

        self.button_frame = tk.Frame(self.root)
        self.button_frame.grid_columnconfigure(0, weight=1)

        self.btn_zurueck = tk.Button(self.button_frame, text="Zurück", command=self.__zurueck, state="disabled")
        self.btn_weiter = tk.Button(self.button_frame, text="Weiter", command=self.__weiter)
        self.btn_uninstall = tk.Button(self.button_frame, text="Uninstall", command=self.__uninstall)


    def _layout_widgets(self):
        self.root.grid_rowconfigure(3, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

        self.label = tk.Label(self.root, text="Installationsoptionen wählen", font=("Arial", 14))
        self.label.grid(row=0, column=0, columnspan=2, padx=20, pady=(20,10), sticky="w")

        self.check1.grid(row=1, column=0, padx=20, sticky="w")
        self.check2.grid(row=2, column=0, padx=20, sticky="w")

        self.button_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=20, padx=20)
        self.btn_zurueck.grid(row=0, column=1, sticky="e", padx=40)
        self.btn_weiter.grid(row=0, column=2, sticky="e")
        self.btn_uninstall.grid(row=0, column=0, sticky="w")


    def __öffen_dataidialog(self):
        self.label3.grid_remove()
        start_dir = r"C:\Program Files"
        pfad = filedialog.askdirectory(initialdir=start_dir)
        if pfad:
            pfad = Path(pfad) / PROGRAM_NAME
            pfad_win = str(PureWindowsPath(pfad))
            self.entry_var.set(pfad_win)


    def __zurueck(self):
        self.entry.grid_remove()
        self.open_btn.grid_remove()

        if self.t is not None:
            threading.Event().set()

        self.label.grid()
        self.label.config(text="Installationsoptionen wählen")
        self.check1.grid()
        self.check2.grid()

        self.btn_zurueck.config(state="disabled")
        self.btn_weiter.config(state="normal")
        self.btn_weiter.config(text="Weiter", command=self.__weiter)
        self.btn_uninstall.grid(row=0, column=0)
        self.btn_weiter.focus_set()

        self.zustand = "start"


    def __weiter(self):
        if self.zustand == "start":
            self.label.config(text="Installationspfad wählen", font=("Arial", 14))
            self.check1.grid_remove()
            self.check2.grid_remove()

            self.entry.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="we")
            self.entry.delete(0, tk.END)
            value = os.path.join(Path("C:\\Program Files"), PROGRAM_NAME)
            self.entry.insert(0, value)
            self.open_btn.grid(row=2, column=0, columnspan=2, padx=20, sticky="w")

            self.btn_zurueck.config(state="normal")
            self.btn_uninstall.grid_forget()
            self.btn_weiter.config(text="Installieren")
            self.btn_weiter.focus_set()

            self.zustand = "installieren"

        elif self.zustand == "installieren":
            self.__installieren()


    def __uninstall(self):
        self.check1.grid_remove()
        self.check2.grid_remove()

        self.btn_weiter.config(text="Uninstall", command=self.__uninstalling)
        self.btn_zurueck.config(state="normal", command=self.__zurueck)
        self.btn_uninstall.grid_forget()

        self.t = threading.Thread(target=self.__hintergrundsuche, daemon=True).start()
        self.label.config(text="Suche nach installiertem Programm")
        self.btn_weiter.config(state="disabled")


    def __hintergrundsuche(self):
        nachricht, pfad = finden()
        self.entry.delete(0, tk.END)

        if pfad:
            text = nachricht
            self.pfad = str(PureWindowsPath(pfad))
            self.entry.insert(0, self.pfad)
            self.entry.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="we")
        else:
            self.label1 = tk.Label(self.root, text=nachricht, font=("Arial", 11), pady=0, justify="left")
            self.label1.grid(row=2, column=0, padx=30, pady=0, sticky="w")
            self.label1.config(text=nachricht)
            msg, found = finden_()
            if pfad:
                self.label1.grid_remove()
                text = msg
                self.pfad =str(PureWindowsPath(found))
                self.entry.insert(0, self.pfad)
                self.entry.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="we")
            else:
                self.label1.grid_remove()
                text = msg
                self.entry.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="we")
                self.open_btn.grid(row=2, column=0, columnspan=2, padx=20, sticky="w")

        if self.pfad.strip():
            self.btn_weiter.config(state="normal")

        self.label.config(text=text)


    def __installieren(self):
        self.open_btn.grid_remove()
        self.entry.grid_remove()
        self.btn_zurueck.grid_remove()
        self.btn_uninstall.destroy()

        self.progressbar.grid(row=2, column=0, columnspan=2, padx=20, pady=10, sticky="we")
        self.progressbar['maximum'] = 100
        self.progressbar['value'] = 0

        self.entry.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="we")
        self.pfad = Path(self.entry_var.get())

        self.t = threading.Thread(target=self.__hintergrundinstall, args=(self.progressbar,), daemon=True).start()
        self.label.config(text=f'Installiere Programm "{PROGRAM_NAME}"')

        self.btn_weiter.grid_forget()


    def __hintergrundinstall(self, progressbar):
        erfolg = False
        fehler = ""

        try:
            erfolg, fehler = install_prog(self.pfad)
            for i in range(101):
                time.sleep(0.05)
                progressbar.after(0, lambda v=i: progressbar.__setitem__('value', v))
                progressbar.update_idletasks()
        except PermissionError:
            if not ctypes.windll.shell32.IsUserAnAdmin():
                self.root.withdraw()  # Hauptfenster ausblenden
                result = messagebox.askokcancel("Fehler, keine Berechtigung.", "Erneut mit Adminrechten starten,\num in den Ordner zu installieren?")

                if result:
                    params = " ".join([f'"{arg}"' for arg in sys.argv])
                    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
                    self.__beenden()
                else:
                    #self.__beenden()   
                    self.root.deiconify()
                    self.progressbar.grid_forget()
                    self.label.config(text="Installationspfad wählen")
                    self.check1.grid_remove()
                    self.check2.grid_remove()
                    self.entry.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="we")
                    self.entry.delete(0, tk.END)
                    self.open_btn.grid(row=2, column=0, columnspan=2, padx=20, sticky="w")
                    self.btn_zurueck.grid(row=0, column=1, sticky="e", padx=40)
                    self.btn_weiter.grid(row=0, column=2, sticky="e")
                    self.label3.grid(row=3, column=0, padx=30, pady=0, sticky="w")
                    self.label3.config(text="Der Pfad C:\\Program Files erforder adminrechte\nanderen Pfad auswählen.")
                    self.btn_zurueck.config(state="disabled")
                    self.btn_uninstall.grid_forget()
                    self.btn_weiter.config(text="Installieren")
                    self.btn_weiter.focus_set()
                    return

            else:
                self.label.config(text=fehler)
                self.__beenden()

        if erfolg:
            self.label.config(text="Installation abgeschlossen")
            if self.var1.get() or self.var2.get():
                self.progressbar.grid_remove()
                self.label1 = tk.Label(self.root, text="Erstelle Verknüpfungen", font=("Arial", 11), pady=0, justify="left")
                self.label1.grid(row=2, column=0, padx=30, pady=0, sticky="w")
            text = ""

            if self.var1.get():
                success, fehler = Menueintrag(self.pfad, ICON)

                if success:
                    text += "Menueintrag erstellt\n"
                else:
                    text += fehler

            if self.var2.get():
                success, fehler = Verknüpfung_erstellen(self.pfad, ICON)

                if success:
                    text += "Desktopverknüpfung erstellt"
                else:
                    text += fehler

            self.label2 = tk.Label(self.root, text=text.strip(), font=("Arial", 11), pady=0, justify="left")
            self.label2.grid(row=3, column=0, padx=30, pady=0, sticky="w")
            self.btn_weiter.grid(row=0, column=2)
            self.btn_weiter.config(text="Fertig", command=self.__beenden)
            self.btn_weiter.focus_set()
        else:
            self.label.config(text="Installation ergab Fehler:")
            self.label = tk.Label(self.root, text=fehler, font=("Arial", 11), pady=0, justify="left")
            self.label.grid(row=2, column=0, padx=30, pady=0, sticky="w")
            #self.label.grid_remove()
            self.btn_weiter.grid(row=0, column=0)
            self.btn_weiter.config(text="Beenden", command=self.__beenden)
            self.btn_weiter.focus_set()


    def __uninstalling(self):
        self.btn_weiter.config(text="Beenden", state="disabled", command=self.__beenden)
        self.open_btn.grid_remove()
        self.entry.grid_remove()
        self.btn_zurueck.destroy()

        erfolg = False
        nachricht = ""

        try:
            erfolg, nachricht = entfernen(self.pfad)
        except PermissionError:
            if not ctypes.windll.shell32.IsUserAnAdmin():
                self.root.withdraw()  # Hauptfenster ausblenden

            result = messagebox.askyesno("Fehler", "Erneut mit Adminrechten starten?")
            if result:
                params = " ".join([f'"{arg}"' for arg in sys.argv])
                ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
                self.__beenden()
            else:
                self.label.config(text=nachricht)
                self.__beenden()

        if erfolg:
            ret = (f"Programm {PROGRAM_NAME} entfernt.")
            self.btn_weiter.config(state="normal")
        else:
            ret = f"Programm {PROGRAM_NAME} nicht entfernt.\n" + nachricht

        self.label.config(text=ret)

        erfolgreich, fehler = Menueintrag_entf()
        text = ""
        if erfolgreich:
            text += fehler + "\n"
        else:
            text += fehler

        success, fehler = Verknüpfung_löschen()

        if success:
            text += fehler + "\n"
        else:
            text += fehler

        self.label1 = tk.Label(self.root, text=text.strip(), font=("Arial", 11), pady=0, justify="left")
        self.label1.grid(row=2, column=0, padx=30, pady=0, sticky="w")


    def __beenden(self):
        self.root.quit()
        self.root.destroy()


    def __xbeenden(self):
        if messagebox.askyesno("Beenden", "Soll die Installation beendet werden?"):
            try:
                pfad = Path(self.pfad)
                if pfad.exists() and PROGRAM_NAME in pfad.name:
                    shutil.rmtree(pfad)

                Verknüpfung_löschen()

                Menueintrag_entf()

            except:
                pass

            self.__beenden()    


    def __ueber(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("Über dieses Programm:")
        about_window.geometry("300x150")
        about_window.resizable(False, False)
        about_window.transient(self.root)  # Fenster bleibt immer über dem Hauptfenster
        about_window.grab_set()

        tk.Label(about_window, text="Wininstaller v2.5.4", font=("Arial", 12, "bold")).pack(pady=10)
        tk.Label(about_window, text="Erstellt von: Ingenieurbuero Ihb. Armin Herzner").pack()
        tk.Label(about_window, text="© 2025, Alle Rechte vorbehalten").pack()
        tk.Button(about_window, text="Schließen", command=about_window.destroy).pack(pady=10)

        self.root.wait_window(about_window)
