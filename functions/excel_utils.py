import os
import csv
from openpyxl import Workbook, load_workbook

# Pour toujours pointer vers Database/station.csv, quel que soit le cwd
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
CSV_PATH = os.path.join(BASE_DIR, "Database", "station.csv")


def load_stations_info():
    """
    Lit station.csv et renvoie { code_station: (commune, lat, lon, z_cc49) }.
    Détecte le séparateur CSV automatiquement.
    """
    station_info = {}
    with open(CSV_PATH, "r", encoding="utf-8-sig") as f:
        sample = f.read(2048)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=";,")
        except csv.Error:
            dialect = csv.get_dialect("excel")
        reader = csv.DictReader(f, dialect=dialect)

        for row in reader:
            code = row.get("Station", "").strip().upper()
            if not code:
                continue

            commune = row.get("Commune", "").strip()
            # Latitude / Longitude
            try:
                lat = row.get("Latitude", "").strip().replace(",", ".")
            except (ValueError, TypeError):
                lat = None
            try:
                lon = row.get("Longitude", "").strip().replace(",", ".")
            except (ValueError, TypeError):
                lon = None

            # Z_CC49 (virgule → point)
            z_raw = row.get("Z_CC49", "").strip().replace(",", ".")
            try:
                z = float(z_raw)
            except (ValueError, TypeError):
                z = None

            station_info[code] = (commune, lat, lon, z)

    # DEBUG : affiche un résumé rapide dans la console
    print(f"[DEBUG] {len(station_info)} stations chargées :", list(station_info.keys())[:10])
    return station_info


def parse_photo_date_time(photo_name):
    """
    Extrait et formate la date/heure depuis le nom de la photo.
    Exemple : "20241025_105000_SE02.jpg" -> "25/10/2024 10h50m00"
    """
    name_no_ext, _ = os.path.splitext(photo_name)
    parts = name_no_ext.split('_')
    if len(parts) >= 2:
        date_str, time_str = parts[0], parts[1]
        if len(date_str) == 8 and len(time_str) == 6:
            formatted_date = f"{date_str[6:8]}/{date_str[4:6]}/{date_str[:4]}"
            formatted_time = f"{time_str[:2]}h{time_str[2:4]}m{time_str[4:]}"
            return f"{formatted_date} {formatted_time}"
    return ""


def create_or_update_excel(input_folder, output_folder):
    """
    Parcourt le dossier d'entrée (contenant les sous-dossiers des stations)
    et crée ou met à jour le fichier Excel dans le dossier de sortie.

    Chaque sous-dossier correspond à une station (ex. "SW47") et aura sa propre feuille
    avec les colonnes: [Nom de la photo, Date / Heure, Résultat].

    La feuille "Résumé" sera reconstruite et contiendra :
      [Commune, Station, Latitude, Longitude, Z_CC49, Nombre de Photos].
    """
    # 1) Charge les infos depuis le CSV
    station_info = load_stations_info()

    # 2) Récupère les dossiers de stations et leurs photos
    stations_data = {}
    for station in os.listdir(input_folder):
        station_path = os.path.join(input_folder, station)
        if os.path.isdir(station_path):
            photos = [
                p for p in os.listdir(station_path)
                if p.lower().endswith(('.jpg', '.jpeg', '.png'))
            ]
            stations_data[station] = photos

    # 3) Affiche les codes manquants dans le CSV
    missing = [s for s in stations_data if s.strip().upper() not in station_info]
    if missing:
        print("[WARN] Stations non trouvées dans station.csv :", missing)

    # 4) Ouvre ou crée le fichier Excel
    file_path = os.path.join(output_folder, "resultats_photos.xlsx")
    if os.path.exists(file_path):
        wb = load_workbook(file_path)
    else:
        wb = Workbook()

    # 5) Prépare la feuille "Résumé"
    if "Résumé" in wb.sheetnames:
        ws_summary = wb["Résumé"]
        ws_summary.delete_rows(2, ws_summary.max_row)
    else:
        ws_summary = wb.active
        ws_summary.title = "Résumé"
        ws_summary.append([
            "Commune", "Station", "Latitude",
            "Longitude", "Z_CC49", "Nombre de Photos"
        ])
    # Réécriture de l’en-tête (toujours, pour l’ordre des colonnes)
    headers = ["Commune", "Station", "Latitude", "Longitude", "Z_CC49", "Nombre de Photos"]
    for col_idx, title in enumerate(headers, start=1):
        ws_summary.cell(row=1, column=col_idx, value=title)

    # 6) Crée/Met à jour chaque feuille de station
    for station, photos in stations_data.items():
        if station in wb.sheetnames:
            ws_station = wb[station]
        else:
            ws_station = wb.create_sheet(station)
            ws_station.append(["Nom de la photo", "Date / Heure", "Résultat"])

        existing = {
            row[0] for row in ws_station.iter_rows(
                min_row=2, max_col=1, values_only=True
            ) if row[0]
        }
        for photo in photos:
            if photo not in existing:
                ws_station.append([photo, parse_photo_date_time(photo), ""])

    # 7) Remplit la feuille "Résumé"
    for idx, station in enumerate(stations_data, start=2):
        upper = station.strip().upper()
        commune = lat = lon = z_cc49 = None
        if upper in station_info:
            commune, lat, lon, z_cc49 = station_info[upper]
        count = wb[station].max_row - 1 if station in wb.sheetnames else 0

        ws_summary.cell(row=idx, column=1, value=commune)
        ws_summary.cell(row=idx, column=2, value=station)
        ws_summary.cell(row=idx, column=3, value=lat)
        ws_summary.cell(row=idx, column=4, value=lon)
        ws_summary.cell(row=idx, column=5, value=z_cc49)
        ws_summary.cell(row=idx, column=6, value=count)

    # 8) Sauvegarde et message retour
    wb.save(file_path)
    msg = "Mise à jour effectuée." if os.path.exists(file_path) else "Nouveau fichier Excel créé."
    return os.path.abspath(file_path), msg


def list_photos_without_result(excel_file):
    """
    Lit le fichier Excel et retourne une liste de tuples
    (sheet_name, row_index, photo_name)
    pour chaque photo dont la colonne 'Résultat' est vide.
    """
    if not os.path.exists(excel_file):
        return []

    wb = load_workbook(excel_file, data_only=True)
    photos_missing = []
    for sheet in wb.sheetnames:
        if sheet == "Résumé":
            continue
        ws = wb[sheet]
        for i in range(2, ws.max_row + 1):
            name = ws.cell(row=i, column=1).value
            res  = ws.cell(row=i, column=3).value
            if name and not str(res).strip():
                photos_missing.append((sheet, i, name))
    return photos_missing


def update_excel_result(excel_file, sheet_name, row_idx, new_value):
    """
    Met à jour la cellule 'Résultat' (colonne 3) de la feuille 'sheet_name'
    à la ligne row_idx dans le fichier Excel, puis sauvegarde.
    """
    wb = load_workbook(excel_file)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"La feuille {sheet_name} n'existe pas dans {excel_file}.")
    ws = wb[sheet_name]
    print(f"Avant mise à jour, ({row_idx},3) = {ws.cell(row=row_idx, column=3).value}")
    ws.cell(row=row_idx, column=3, value=new_value)
    print(f"Après mise à jour, ({row_idx},3) = {ws.cell(row=row_idx, column=3).value}")
    wb.save(excel_file)
