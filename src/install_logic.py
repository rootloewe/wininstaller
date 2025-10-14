import string
import os
from pathlib import Path, PureWindowsPath
from win32com.client import Dispatch
import zipfile
import shutil
import pythoncom

PROGRAM_NAME = "einlesen"
ICON = r"C:\Users\armin\Desktop\projekte\wininstaller\icon.ico"
ZIP = r"main.zip"


def install_prog(pfad):
    #extract_path = str(PureWindowsPath(pfad))

    # if os.path.basename(pfad_win).lower() != PROGRAM_NAME.lower():
    #     extract_path = os.path.join(pfad_win, PROGRAM_NAME)
    # else:
    #     extract_path = pfad_win
    print(pfad)
    if not pfad.exists():
        pfad.mkdir(parents=True, exist_ok=True)

    try:
        zip_ref = zipfile.ZipFile(ZIP, 'r')
        zip_ref.extractall(pfad)
        zip_ref.close()

        print("ZIP-Datei entpackt und Inhalt kopiert.")

    except zipfile.BadZipFile:
        print("Fehler: Die ZIP-Datei ist beschädigt oder kein gültiges ZIP-Format.")
        return False, "Die ZIP-Datei ist beschädigt oder kein gültiges ZIP-Format."
    except FileNotFoundError:
        print("Fehler: Die ZIP-Datei wurde nicht gefunden.")
        return False, "Die ZIP-Datei wurde nicht gefunden."
    except PermissionError:
        print("Keine Amdinrechte")
        raise
    except Exception as e:
        print(f"Unbekannter Fehler {e}")

    return True, ""


def Verknüpfung_erstellen(target, icon=None):
    shortcut_path =  str(Path.home() / "Desktop" / f"{PROGRAM_NAME}.lnk")
    pythoncom.CoInitialize()
    try:
        shell = Dispatch('WScript.Shell')
        datei = os.path.join(str(target), f"{PROGRAM_NAME}.exe")
        shortcut = shell.CreateShortcut(shortcut_path)
        shortcut.TargetPath = datei
        shortcut.WorkingDirectory = os.path.dirname(datei)
        shortcut.IconLocation = icon if icon and os.path.exists(icon) else datei
        shortcut.save()
        print(f"Verknüpfung erfolgreich erstellt: {shortcut_path}\n")
    except Exception as e:
        print(f"Fehler beim Erstellen der Verknüpfung: {e}")
        return False, (f"Fehler beim Erstellen der Verknüpfung:\n {e}\n")
    finally:
        pythoncom.CoUninitialize()

    return True, ""


def Menueintrag(pfad, icon=None):
    pfad = Path(pfad) / PROGRAM_NAME
    pfad_win = str(PureWindowsPath(pfad))
    pythoncom.CoInitialize()

    try:
        start_menu = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs")

        menu_pfad = os.path.join(start_menu, f"{PROGRAM_NAME}.lnk")

        if os.path.exists(menu_pfad):
            os.makedirs(menu_pfad)

        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortcut(menu_pfad)
        shortcut.TargetPath = pfad_win
        shortcut.WorkingDirectory = os.path.dirname(pfad_win)

        if icon and os.path.exists(icon):
            shortcut.IconLocation = icon
        else:
            shortcut.IconLocation = pfad_win

        shortcut.save()
        print("Menueintrag erfolgreich erstellt\n")

    except Exception as e:
        print(f"Fehler beim Erstellen der Startmenü-Verknüpfung:\n {e}")
        return False, f"Fehler Startmenü-Eintrag:\n{e}\n"

    finally:
        pythoncom.CoUninitialize()

    return True, ""


def finden():
    gesuchter_dateiname = os.path.join(f"{PROGRAM_NAME}.exe")

    startverzeichnis = Path(r"C:\\")

    for datei in startverzeichnis.rglob(gesuchter_dateiname, case_sensitive=True):
        if datei.is_file():
            print(datei.resolve())
            return  "Programm gefunden", datei.resolve()

    return "Programm nicht gefunden\nerweitere Suche über andere Laufwerke.", ""


def finden_():
    gesuchter_dateiname = os.path.join(f"{PROGRAM_NAME}.exe")
    laufwerke = [f"{buchstabe}:\\" for buchstabe in string.ascii_uppercase if os.path.exists(f"{buchstabe}:\\")]

    for laufwerk in laufwerke:
        pfad = Path(laufwerk)
        for datei in pfad.rglob(gesuchter_dateiname, case_sensitive=True):
            if datei.is_file():
                print(datei.resolve())
                return f"Entferne Programm {PROGRAM_NAME}:", datei.resolve()

    return "Nichts gefunden, gib den Pfad an:", ""


def entfernen(pfad):
    if pfad and pfad is not None:
        pfad = os.path.dirname(pfad)
        try:
            shutil.rmtree(pfad)
            print(f"Pfad gelöscht: {pfad}")
            return True, ""
        except PermissionError:
            print(f"Keine Berechtigung zum Löschen vom Pfad: {pfad}")
            raise
        except Exception as e:
            print(f"Fehler beim Löschen von {pfad}: {e}")
            return False, (f"Fehler beim Löschen von {pfad}:\n {e}")
    else:
        print(f"Pfad nicht gefunden (plötzlich gelöscht?):\n {pfad}")
        return False, (f"Pfad nicht gefunden (plötzlich gelöscht?):\n {pfad}")


def Verknüpfung_löschen():
    datei_name = Path.home() / "Desktop" / f"{PROGRAM_NAME}.lnk"

    try:
        if datei_name.exists():
            datei_name.unlink()
            print(f"Datei {datei_name} wurde gelöscht.")
            return True, "Desktopverknüpfung entfernt.\n"
        else:
            print("Datei existiert nicht.")
            return False, "Keine Desktopverknüpfung gefunden.\n"
    except Exception as e:
        print(f"Fehler beim Löschen der Datei: {e}")
        return False, f"Fehler beim Löschen der Datei:\n {e}"


def Menueintrag_entf():
    start_menu2 = os.path.join(
        os.path.expanduser("~"),
        "AppData", "Roaming", "Microsoft", "Windows",
        "Start Menu",
        "Programs"
    )
    shortcut_path2 = os.path.join(start_menu2, f"{PROGRAM_NAME}.lnk")

    try:
        if os.path.exists(shortcut_path2):
            os.remove(shortcut_path2)
            print("Menueintrag wurde entfernt.\n")
            return True, "Menueintrag entfernt"
        else:
            print("Menaueintrag nicht vorhanden\n")
            return False, "Menaueintrag nicht vorhanden\n"

    except Exception as e:
        print(f"Fehler beim Entfernen der Verknüpfung: {e}")
        return False, f"Fehler beim Entfernen der Verknüpfung:\n {e}"
