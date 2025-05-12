import csv
import os

def load_stations_info():
    """
    Lit station.csv et renvoie { code_station: (commune, lat, lon, z_cc49) }.
    Détecte automatiquement le séparateur (; ou ,).
    """
    station_info = {}
    csv_path = os.path.join("", "station.csv")

    with open(csv_path, "r", encoding="utf-8-sig") as f:
        sample = f.read(2048)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=";,")
        except csv.Error:
            dialect = csv.get_dialect('excel')
        reader = csv.DictReader(f, dialect=dialect)

        for row in reader:
            code = row.get("Station", "").strip()
            if not code:
                continue

            commune = row.get("Commune", "").strip()
            # lat / lon
            try:
                lat = float(row.get("Latitude", "").strip())
            except ValueError:
                lat = None
            try:
                lon = float(row.get("Longitude", "").strip())
            except ValueError:
                lon = None
            # Z_CC49 (virgule → point)
            z_raw = row.get("Z_CC49", "").strip()
            try:
                z = float(z_raw.replace(",", "."))
            except ValueError:
                z = None

            station_info[code] = (commune, lat, lon, z)

    # DEBUG: afficher un échantillon
    print(f"[load_stations_info] {len(station_info)} stations chargées, quelques-unes : "
          f"{list(station_info.keys())[:5]}")

    return station_info
