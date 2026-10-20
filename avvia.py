#!/usr/bin/env python3
"""
avvia.py - aggiorna i dati e apre l'app nel browser (doppio clic su Avvia_Bandi_Veneto.bat / .command / .sh).

  python avvia.py                 aggiornamento normale (circa 3-6 minuti: schede degli atti aperti + nuovi chiusi)
  python avvia.py --veloce        solo liste e calendari (circa 20 secondi)
  python avvia.py --completo      passata lunga: recupera molte schede di dettaglio storiche (fino a 30 minuti)
  python avvia.py --solo-apri     non aggiorna, apre soltanto l'app
Se l'aggiornamento fallisce (es. portale non raggiungibile) l'app si apre comunque con gli ultimi dati salvati.
"""
import argparse, os, subprocess, sys, webbrowser
from pathlib import Path

HERE = Path(__file__).resolve().parent


def dipendenze():
    mancanti = []
    for mod, pkg in (("requests", "requests"), ("bs4", "beautifulsoup4"), ("openpyxl", "openpyxl")):
        try:
            __import__(mod)
        except ImportError:
            mancanti.append(pkg)
    if mancanti:
        print("Installo i componenti mancanti:", ", ".join(mancanti))
        r = subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", *mancanti])
        if r.returncode != 0:
            print("Installazione non riuscita. Prova a mano: pip install -r requirements.txt")
            return False
    return True


def main():
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--veloce", action="store_true")
    g.add_argument("--completo", action="store_true")
    g.add_argument("--solo-apri", action="store_true")
    ap.add_argument("--no-apri", action="store_true", help="non aprire il browser (per prove)")
    a = ap.parse_args()

    if sys.version_info < (3, 9):
        print("Serve Python 3.9 o successivo (https://www.python.org/downloads/).")
        return 1
    ok = True
    if not a.solo_apri:
        if not dipendenze():
            ok = False
        else:
            if a.veloce:
                extra = ["--no-refresh-open", "--max-details", "0"]
            elif a.completo:
                extra = ["--max-details", "400", "--time-budget", "1800"]
            else:
                extra = ["--max-details", "60", "--time-budget", "300"]
            print("Aggiornamento in corso... (puoi lasciare la finestra aperta)")
            r = subprocess.run([sys.executable, str(HERE / "aggiorna_bandi.py"), *extra], cwd=str(HERE))
            ok = r.returncode == 0
            print("Aggiornamento completato." if ok else
                  "Aggiornamento NON riuscito (portale non raggiungibile?): apro l'app con gli ultimi dati salvati.")
    html = HERE / "bandi_veneto.html"
    if not html.exists():
        print("File bandi_veneto.html non trovato.")
        return 1
    if not a.no_apri:
        webbrowser.open(html.as_uri())
        print("App aperta nel browser:", html)
    return 0 if ok else 2


if __name__ == "__main__":
    sys.exit(main())
