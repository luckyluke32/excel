import os
import time
from typing import List, Optional

import requests
from openpyxl import Workbook

OVERPASS_URL = "https://overpass-api.de/api/interpreter"
ATTICA_RELATION_ID = 4477553
ATTICA_AREA_ID = 3600000000 + ATTICA_RELATION_ID
NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
FOURSQUARE_URL = "https://api.foursquare.com/v3/places/search"


def get_attica_municipalities() -> List[str]:
    """Retrieve a list of Attica municipalities from OpenStreetMap."""
    query = f"[out:json];area({ATTICA_AREA_ID})->.a;relation[admin_level=7](area.a);out tags;"
    res = requests.post(OVERPASS_URL, data={"data": query}, timeout=120)
    res.raise_for_status()
    data = res.json()
    names = {
        el["tags"].get("name:en", el["tags"].get("name"))
        for el in data.get("elements", [])
        if el.get("tags")
    }
    return sorted(names)


def get_bbox(municipality: str) -> Optional[dict]:
    """Return bounding box for the municipality using Nominatim."""
    q = f"{municipality}, Attica, Greece"
    params = {"q": q, "format": "json", "limit": 1}
    headers = {"User-Agent": "attica-script"}
    res = requests.get(NOMINATIM_URL, params=params, headers=headers, timeout=60)
    res.raise_for_status()
    data = res.json()
    if not data:
        return None
    bb = data[0].get("boundingbox", [])
    if len(bb) != 4:
        return None
    south, north, west, east = map(float, bb)
    return {"south": south, "north": north, "west": west, "east": east}


def fetch_restaurants(municipality: str, api_key: str) -> List[dict]:
    """Fetch restaurant names and emails for a municipality via Foursquare."""
    bbox = get_bbox(municipality)
    if not bbox:
        return []
    headers = {"Authorization": api_key}
    params = {
        "categories": "13065",
        "fields": "name,contacts",
        "limit": 50,
        "sw": f"{bbox['south']},{bbox['west']}",
        "ne": f"{bbox['north']},{bbox['east']}",
    }
    restaurants = []
    while True:
        res = requests.get(FOURSQUARE_URL, headers=headers, params=params, timeout=60)
        res.raise_for_status()
        data = res.json()
        for place in data.get("results", []):
            name = place.get("name")
            email = place.get("contacts", {}).get("email")
            if name and email:
                restaurants.append({"name": name, "email": email})
        cursor = data.get("context", {}).get("next_cursor")
        if cursor:
            params["cursor"] = cursor
            time.sleep(1)
        else:
            break
    return restaurants


def main(output_file: str = "attica_restaurants.xlsx") -> None:
    api_key = os.environ.get("FOURSQUARE_API_KEY") or os.environ.get("FSQ_API_KEY")
    if not api_key:
        raise RuntimeError("FOURSQUARE_API_KEY environment variable not set")

    wb = Workbook()
    first_sheet = True
    municipalities = get_attica_municipalities()
    for municipality in municipalities:
        if first_sheet:
            ws = wb.active
            ws.title = municipality
            first_sheet = False
        else:
            ws = wb.create_sheet(title=municipality)
        ws.append(["Restaurant Name", "Email"])
        restaurants = fetch_restaurants(municipality, api_key)
        for rest in restaurants:
            ws.append([rest.get("name", ""), rest.get("email", "")])
    wb.save(output_file)
    print(f"Excel file '{output_file}' created.")


if __name__ == "__main__":
    main()
