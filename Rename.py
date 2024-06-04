# v3.4


import os
import pandas as pd  # type: ignore
import glob
import sys
import datetime
import signal

clear = lambda: os.system('cls')

def signal_handler(sig, frame):
    log_message("Program megszakítva.")
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)


# A mappa mindig a fájlt tartalmazó mappa.
current_directory = os.path.dirname(os.path.abspath(__file__))
folderName = os.path.basename(current_directory)

log_file_path = os.path.join(current_directory, 'log')




# Logolás
def log_message(message):
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        log_file.write(f"{datetime.datetime.now()}: {message}\n")
        os.system("attrib +h log")

    print(message)



# Érvénytelen karakterek lecserélése
def sanitize_filename(value):
    return value.replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('\n',' ')



def rename_files():
    clear()
    renamed = 0
    error = 0
    if os.path.exists(log_file_path):
        with open(log_file_path, 'r', encoding='utf-8') as log_file:
            lines = log_file.readlines()
            for line in lines:
                if 'Módosítva: ' in line:
                    clear()
                    print("A fájlok már egyszer módosítva lettek a(z) " + folderName + " mappában,\n"
                          + "így az újbóli átnevezés nem ajánlott, főleg ha hibák is történtek.\n\n"
                          + "Amennyiben kézileg történt módosítás, visszaállítás után újból át szeretnéd nevezni, vagy egyéb okból\n"
                          + "úgy döntesz, hogy mégis szeretnéd az átnevezést újból végrehajtani,\n"
                          + "akkor írd be hogy 'Megerősítem', majd indítsd el a programot újból.\n"
                          + "Ekkor azonban végleg elvész az előző módosítások visszaállításának lehetősége.\n\n\n\n"
                          + "Megerősíted, hogy újból módosítani szeretnéd a mappában lévő fájlokat,\n"
                          + "még akkor is, ha az előző visszaállítás lehetőségét elveszíted?\n")
                    
                    response = input()
                    if response != "Megerősítem":
                        print("Nincs megerősítve. Módosítás nem történt, a program kilép.")
                        sys.exit()
                    else:
                        print("Megerősítve. A program kilép, indítsd újra a programot.")
                        log_file.close()
                        os.remove(log_file_path)
                        sys.exit()
                

    # Keressük az Excel fájlt ami alapján az elnevezések történnek
    excel_files = glob.glob(os.path.join(current_directory, '*_Filelist.xlsx')) + glob.glob(os.path.join(current_directory, '*_Filelist.xlsm'))
    if not excel_files:
        print("Nem található *_Filelist nevű Excel fájl a mappában! Kérlek hozd létre, és/vagy nevezd el.")
        sys.exit()
    elif len(excel_files) > 1:
        print("Több *_Filelist nevű Excel fájl található a mappában! Kérlek egyszerre csak egy legyen.\nAmennyiben csak egy van, de az meg van nyitva, kérlek zárd be.")
        sys.exit()
    else:
        excel_file = excel_files[0]
        excelName = os.path.basename(excel_file)
        log_message("A program a(z) " + folderName + " mappában lévő fájlok átnevezésére készül.\n"
                    + "Ehhez a(z) " + excelName + " fájlt fogja használni.\n\n"
                    + "Bár az esetleges hibás módosításokat (amiket a program hajt végre)\n"
                    + "vissza lehet állítani a -v vagy --vissza kapcsolók megadásával,\n"
                    + "előfordulhat, hogy valahány hibát kézileg lehet majd csak visszaállítani.\n\n\n\n"
                    + "Ezek tudatában mehet tovább a program? [I/N]")


        # Felhasználói válasz bekérése
        user_response = input().strip().upper()
        if user_response != 'I':
            print("A program kilép. Módosítás nem történt.")
            os.remove("log")
            sys.exit()


    # Az Excel fájl betöltése
    df = pd.read_excel(excel_file)

    # Ellenőrzések
    if df.empty:
        clear()
        log_message("Hiba: Az Excel fájl üres.\n\nNem történt átnevezés.")
        sys.exit()

    if df.shape[1] < 3:
        clear()
        log_message("Hiba: Az Excel fájlnak legalább három oszlopot kell tartalmaznia.\nKérlek ellenőrizd, hogy az Excel fájl a mintának megfelel.\n\nNem történt átnevezés.")
        sys.exit()

    if df.iloc[:, 0].isnull().all():
        clear()
        log_message("Hiba: Az első oszlop üres.\nKérlek ellenőrizd, hogy az Excel fájl a mintának megfelel.\n\nNem történt átnevezés.")
        sys.exit()

    if df.iloc[:, 1].isnull().all():
        clear()
        log_message("Hiba: A második oszlop üres.\nKérlek ellenőrizd, hogy az Excel fájl a mintának megfelel.\n\nNem történt átnevezés.")
        sys.exit()

    if df.iloc[:, 2].isnull().all():
        clear()
        log_message("Hiba: A harmadik oszlop üres.\nKérlek ellenőrizd, hogy az Excel fájl a mintának megfelel.\n\nNem történt átnevezés.")
        sys.exit()



    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        log_file.write("\n\n\n\n\t\t\t\t\t\t\t\t\t\t\t\t######## ! INNENTŐL SEMMIKÉPP NE MÓDOSÍTSD ! ########\n\n"
                       + "\t\t\t\t\t\t\t\t\t\t\t\tA FÁJL TÁJÉKOZTAT, DE A HELYES MŰKÖDÉST IS BIZTOSÍTJA\n\t\t\t\t\t\t\t\t\t\t\t\tMÓDOSÍTÁS ESETÉN NEM GARANTÁLHATÓ A MEGFELELŐ MŰKÖDÉS"
                       + "\n\n\t\t\t\t\t\t\t\t\t\t\t\t######## ! INNENTŐL SEMMIKÉPP NE MÓDOSÍTSD ! ########\n\n\n\n")
    # Az oszlopok értékei
    file_names = df.iloc[:, 0].tolist()
    prefix_values = df.iloc[:, 2].tolist()  # 3. oszlop értékei
    suffix_values = df.iloc[:, 1].tolist()  # 2. oszlop értékei



    # Minden fájl feldolgozása az Excel táblázat alapján
    for original_name, prefix, suffix in zip(file_names, prefix_values, suffix_values):

        # Az új fájlnév létrehozása
        sanitized_prefix = sanitize_filename(prefix)
        sanitized_suffix = sanitize_filename(suffix)
        name, ext = os.path.splitext(original_name)
        new_filename = f"{sanitized_prefix}___{name}___{sanitized_suffix}{ext}"
            
        # Fájl átnevezése
        old_file = os.path.join(current_directory, original_name)
        new_file = os.path.join(current_directory, new_filename)
            
            
        # Ellenőrizzük, hogy az eredeti fájl létezik-e
        if os.path.exists(old_file):
            os.rename(old_file, new_file)
            log_message(f"Átnevezve: {original_name} -> {new_filename}")
            renamed = renamed + 1
        else:
            print(f"\n\nHiba: {original_name} nem található a mappában!\nEllenőrizd az Excel fájlt, a hibásnak vélt fájl meglétét, illetve a speciális karakterek miatti eltéréseket!\n\n")
            error = error + 1

    if error == 0:
        print(f"\n\nA fájlok átnevezése sikeresen befejeződött. {renamed} fájl módosult.\nHiba nem történt.")
        with open(log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write("Módosítva: Igen\n")
    else:
        print(f"\n\nFájlok átnevezése befejezve. {renamed} fájl módosult.\nAz átnevezés során {error} hiba történt.")
        with open(log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write(f"Módosítva: Részlegesen, {error} hibával.\n")



def restore_files():
    restored = 0
    error = 0
    if not os.path.exists(log_file_path) or os.path.getsize(log_file_path) == 0:
        print("\n\nVisszaállítás nem lehetséges. Okok lehetnek:\n-Nem történt módosítás a program által\n-Meg lett erősítve az újbóli átnevezés.\n\n")
        sys.exit()


    with open(log_file_path, 'r', encoding='utf-8') as log_file:
        lines = log_file.readlines()
        for line in lines:
            if 'Visszaállítva: Igen' in line:
                clear()
                print("A fájlok már vissza lettek állítva korábban, a(z) " + folderName + " mappában az eredeti fájlok vannak.")
                return
    


    clear()
    print("Szeretnéd visszaállítani az eredeti neveket? [I/N]")
    user_response = input().strip().upper()
    if user_response != 'I':
        print("Nem történt módosítás. A program kilép.")
        sys.exit()


    log_message("Eredeti fájlnevek visszaállítása...")
    with open(log_file_path, 'r', encoding='utf-8') as log_file:
        lines = log_file.readlines()
        for line in lines:
            if 'Átnevezve: ' in line:
                parts = line.strip().split('Átnevezve: ')[1].split(' -> ')
                original_name = parts[0]
                new_filename = parts[1]
                old_file = os.path.join(current_directory, new_filename)
                new_file = os.path.join(current_directory, original_name)
                
                if os.path.exists(old_file):
                    os.rename(old_file, new_file)
                    restored = restored + 1
                    log_message(f"Visszaállítva: {new_filename} -> {original_name}")
                else:
                    log_message(f"Hiba: {new_filename} nem található a mappában! Nem sikerült visszaállítani.")
                    error = error + 1


    if error == 0:
        log_message(f"\n\nVisszaállítás sikeres. {restored} fájl visszaállítva.\nHiba nem történt.")
        with open(log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write("Visszaállítva: Igen\n")
    else:
        log_message(f"\n\nVisszaállítás befejezve. {restored} fájl visszaállítva.\nA visszaállítás során {error} hiba történt.")
        with open(log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write(f"Visszaállítva: Részlegesen, {error} hibával.\n")




# A program indítása
if __name__ == "__main__":
    if '--vissza' in sys.argv or '-v' in sys.argv:
        restore_files()
    else:
        rename_files()
