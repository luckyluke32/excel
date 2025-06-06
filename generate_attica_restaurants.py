from openpyxl import Workbook
import requests

OVERPASS_URL = "https://overpass-api.de/api/interpreter"
ATTICA_RELATION_ID = 4477553  # Relation id for the Attica region
ATTICA_AREA_ID = 3600000000 + ATTICA_RELATION_ID


def get_attica_municipalities():
    """Retrieve the list of municipalities within Attica from OpenStreetMap."""
    query = f"[out:json];area({ATTICA_AREA_ID})->.a;relation[admin_level=7](area.a);out tags;"
    res = requests.post(OVERPASS_URL, data={'data': query}, timeout=120)
    res.raise_for_status()
    data = res.json()
    names = {
        el['tags'].get('name:en', el['tags'].get('name'))
        for el in data.get('elements', [])
        if el.get('tags')
    }
    return sorted(names)


def fetch_restaurants(municipality):
    """Fetch restaurants with email information for a municipality using OSM."""

    area_query = f'area["name:en"="{municipality}"][admin_level=7]'
    query = (
        f'[out:json][timeout:120];{area_query}->.a;'
        '(node(area.a)["amenity"="restaurant"]["email"];'
        ' node(area.a)["amenity"="restaurant"]["contact:email"];'
        ' way(area.a)["amenity"="restaurant"]["email"];'
        ' way(area.a)["amenity"="restaurant"]["contact:email"];'
        ');out;'
    )

    res = requests.post(OVERPASS_URL, data={'data': query}, timeout=120)
    res.raise_for_status()
    data = res.json()
    restaurants = []
    for el in data.get('elements', []):
        tags = el.get('tags', {})
        name = tags.get('name')
        email = tags.get('email') or tags.get('contact:email')
        if name and email:
            restaurants.append({'name': name, 'email': email})
    return restaurants


def main(output_file="attica_restaurants.xlsx"):
    """Generate an Excel workbook with one sheet per municipality."""
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

        # Column headers
        ws.append(["Restaurant Name", "Email"])

        restaurants = fetch_restaurants(municipality)

        for rest in restaurants:
            ws.append([rest.get("name", ""), rest.get("email", "")])

    wb.save(output_file)
    print(f"Excel file '{output_file}' created.")


if __name__ == "__main__":
    main()
