"""
FIPE Table Enricher for Auction Lot Spreadsheets
=================================================
Reads an Excel file with vehicle auction lots, queries the FIPE API
for each vehicle's valuation, and writes the result back to the spreadsheet.

Usage:
    python fipe_enricher.py lotes_leilao.xlsx

Output:
    lotes_leilao_com_fipe.xlsx (enriched file with FIPE values)

Requirements:
    pip install openpyxl requests
"""

import sys
import time
import json
import re
import os
from difflib import SequenceMatcher
import requests
import openpyxl

API_BASE = "https://parallelum.com.br/fipe/api/v1"
CACHE_FILE = "fipe_cache.json"

# Brand mapping: spreadsheet name -> (vehicle_type, fipe_brand_code)
# vehicle_type: "carros" or "motos"
BRAND_MAP = {
    # Cars
    "CHEVROLET": ("carros", "23"),
    "CHEV": ("carros", "23"),
    "GM": ("carros", "23"),
    "FIAT": ("carros", "21"),
    "FORD": ("carros", "22"),
    "VW": ("carros", "59"),
    "PEUGEOT": ("carros", "44"),
    "RENAULT": ("carros", "48"),
    "CITROEN": ("carros", "13"),
    "M.BENZ": ("carros", "39"),
    "MERCEDES": ("carros", "39"),
    "TOYOTA": ("carros", "56"),
    "HYUNDAI": ("carros", "26"),
    "CHERY": ("carros", "182"),
    "GEELY": ("carros", "199"),
    # Motorcycles
    "YAMAHA": ("motos", "101"),
    "SUNDOWN": ("motos", "98"),
    "DAFRA": ("motos", "145"),
    "SHINERAY": ("motos", "134"),
    "JTZ": ("motos", "209"),
}

# Honda appears in both cars and motos - detect by model keywords
MOTO_KEYWORDS = [
    "CG", "CBX", "BIZ", "NXR", "XRE", "CB", "TITAN", "FAN", "FACTOR",
    "FAZER", "YBR", "NEO", "LANDER", "CRYPTON", "LEAD", "PCX", "POP",
    "BROS", "TWISTER", "TORNADO", "FALCON", "SPEED", "PHOENIX", "CHOPPER",
    "CARGO", "START", "CROSSER", "XTZ", "FZ25", "BURGMAN", "INTRUDER"
]

CAR_KEYWORDS = [
    "CIVIC", "FIT", "CITY", "HR-V", "HRV", "WR-V", "ACCORD", "CR-V",
    "ML", "CLASS", "TUCSON", "CRETA", "HB20", "IX35"
]


def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def api_get(url, cache, retries=3):
    if url in cache:
        return cache[url]
    for attempt in range(retries):
        try:
            time.sleep(0.5)
            resp = requests.get(url, timeout=15)
            if resp.status_code == 200:
                data = resp.json()
                if isinstance(data, dict) and "error" in data:
                    cache[url] = None
                    return None
                cache[url] = data
                return data
            elif resp.status_code == 429:
                wait = 5 * (attempt + 1)
                print(f"    Rate limited, waiting {wait}s...")
                time.sleep(wait)
            else:
                cache[url] = None
                return None
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(2)
            else:
                print(f"    API error: {e}")
                cache[url] = None
                return None
    return None


def parse_vehicle(description, fab_mod):
    """Parse brand, model, and year from spreadsheet fields."""
    desc = description.strip()

    # Remove I/ or IMP/ prefix (imported vehicles)
    if desc.startswith("I/"):
        desc = desc[2:]
    elif desc.startswith("IMP/"):
        desc = desc[4:]

    # Split brand/model
    if "/" in desc:
        brand = desc.split("/")[0].strip()
        model = desc.split("/", 1)[1].strip()
    else:
        # No slash - try to split by space (e.g. "HONDA C100 BIZ")
        parts = desc.split()
        brand = parts[0]
        model = " ".join(parts[1:]) if len(parts) > 1 else ""

    # Parse year (use model year = second value in "fab/mod")
    year = None
    if fab_mod:
        fab_mod_str = str(fab_mod).strip()
        parts = fab_mod_str.split("/")
        year_str = parts[-1].strip()
        try:
            year = int(year_str)
        except ValueError:
            pass

    return brand.upper(), model.upper(), year


def detect_vehicle_type(brand, model):
    """Determine if a vehicle is a car or motorcycle."""
    brand_upper = brand.upper()

    # Check if brand is in our map with a definitive type
    if brand_upper in BRAND_MAP:
        return BRAND_MAP[brand_upper]

    # Honda: check model keywords
    if brand_upper == "HONDA":
        model_upper = model.upper()
        for kw in MOTO_KEYWORDS:
            if kw in model_upper:
                return ("motos", "80")
        for kw in CAR_KEYWORDS:
            if kw in model_upper:
                return ("carros", "25")
        # Default Honda to moto (more common in Brazil auctions)
        return ("motos", "80")

    # Unknown brand - try both
    return None


def fuzzy_match(needle, haystack, threshold=0.45):
    """Find the best fuzzy match for needle in a list of {codigo, nome} dicts."""
    needle_clean = re.sub(r'[^a-z0-9 ]', '', needle.lower()).strip()
    needle_words = set(needle_clean.split())

    best_score = 0
    best_match = None

    for item in haystack:
        nome = item["nome"]
        nome_clean = re.sub(r'[^a-z0-9 ]', '', nome.lower()).strip()

        # Sequence matcher score
        seq_score = SequenceMatcher(None, needle_clean, nome_clean).ratio()

        # Word overlap bonus
        nome_words = set(nome_clean.split())
        common = needle_words & nome_words
        word_score = len(common) / max(len(needle_words), 1) if needle_words else 0

        # Combined score
        score = seq_score * 0.6 + word_score * 0.4

        if score > best_score:
            best_score = score
            best_match = item

    if best_score >= threshold:
        return best_match, best_score
    return None, 0


def find_fipe_value(brand, model, year, cache):
    """Query FIPE API to find the vehicle value."""
    vtype_info = detect_vehicle_type(brand, model)

    search_types = []
    if vtype_info:
        search_types.append(vtype_info)
    else:
        search_types.append(("carros", None))
        search_types.append(("motos", None))

    for vtype, brand_code in search_types:
        if not brand_code:
            brands_url = f"{API_BASE}/{vtype}/marcas"
            brands_data = api_get(brands_url, cache)
            if not brands_data:
                continue
            brand_match, _ = fuzzy_match(brand, brands_data, 0.5)
            if not brand_match:
                continue
            brand_code = brand_match["codigo"]

        # Get models
        models_url = f"{API_BASE}/{vtype}/marcas/{brand_code}/modelos"
        models_data = api_get(models_url, cache)
        if not models_data or "modelos" not in models_data:
            continue

        models_list = models_data["modelos"]
        model_match, score = fuzzy_match(model, models_list, 0.35)
        if not model_match:
            continue

        model_code = model_match["codigo"]
        model_name = model_match["nome"]

        # Get available years
        anos_url = f"{API_BASE}/{vtype}/marcas/{brand_code}/modelos/{model_code}/anos"
        anos_data = api_get(anos_url, cache)
        if not anos_data:
            continue

        # Find matching year
        year_code = None
        if year:
            for ano in anos_data:
                ano_year = ano["nome"].split()[0] if " " in ano["nome"] else ano["nome"]
                try:
                    if int(ano_year) == year:
                        year_code = ano["codigo"]
                        break
                except ValueError:
                    continue

            # If exact year not found, try closest year
            if not year_code:
                closest = None
                closest_diff = 999
                for ano in anos_data:
                    ano_year_str = ano["nome"].split()[0] if " " in ano["nome"] else ano["nome"]
                    try:
                        ano_year = int(ano_year_str)
                        diff = abs(ano_year - year)
                        if diff < closest_diff:
                            closest_diff = diff
                            closest = ano
                    except ValueError:
                        continue
                if closest and closest_diff <= 2:
                    year_code = closest["codigo"]

        if not year_code and anos_data:
            year_code = anos_data[0]["codigo"]

        if not year_code:
            continue

        # Get price
        price_url = f"{API_BASE}/{vtype}/marcas/{brand_code}/modelos/{model_code}/anos/{year_code}"
        price_data = api_get(price_url, cache)
        if price_data and "Valor" in price_data:
            return {
                "valor": price_data["Valor"],
                "modelo_fipe": price_data.get("Modelo", model_name),
                "ano": price_data.get("AnoModelo", year),
                "combustivel": price_data.get("Combustivel", ""),
                "codigo_fipe": price_data.get("CodigoFipe", ""),
                "referencia": price_data.get("MesReferencia", ""),
                "match_score": round(score, 2),
            }

    return None


def main():
    if len(sys.argv) < 2:
        print("Usage: python fipe_enricher.py <input_file.xlsx>")
        sys.exit(1)

    input_file = sys.argv[1]
    base_name = os.path.splitext(input_file)[0]
    output_file = f"{base_name}_com_fipe.xlsx"

    print(f"Reading {input_file}...")
    wb = openpyxl.load_workbook(input_file)
    ws = wb.active

    total_rows = ws.max_row - 1
    print(f"Found {total_rows} vehicles to process")

    cache = load_cache()

    found = 0
    not_found = 0
    errors = []

    fipe_col = 6  # Column F = "Valor tabela Fipe"

    for row in range(2, ws.max_row + 1):
        lote = ws.cell(row, 1).value
        description = ws.cell(row, 2).value
        fab_mod = ws.cell(row, 4).value

        if not description:
            continue

        brand, model, year = parse_vehicle(description, fab_mod)
        progress = f"[{row-1}/{total_rows}]"
        print(f"{progress} Lote {lote}: {brand} / {model} / {year}", end=" ... ")

        try:
            result = find_fipe_value(brand, model, year, cache)
            if result:
                ws.cell(row, fipe_col).value = result["valor"]
                found += 1
                print(f"OK -> {result['valor']} (FIPE: {result['modelo_fipe']}, match: {result['match_score']})")
            else:
                ws.cell(row, fipe_col).value = "NAO ENCONTRADO"
                not_found += 1
                print("NOT FOUND")
                errors.append(f"Lote {lote}: {description} ({fab_mod})")
        except Exception as e:
            ws.cell(row, fipe_col).value = f"ERRO: {str(e)[:50]}"
            not_found += 1
            print(f"ERROR: {e}")
            errors.append(f"Lote {lote}: {description} - ERROR: {e}")

        # Save cache periodically
        if (row - 1) % 20 == 0:
            save_cache(cache)

    save_cache(cache)

    print(f"\nSaving to {output_file}...")
    wb.save(output_file)

    print(f"\n{'='*50}")
    print(f"RESULTS:")
    print(f"  Total vehicles: {total_rows}")
    print(f"  Found FIPE value: {found}")
    print(f"  Not found: {not_found}")
    print(f"  Success rate: {found/total_rows*100:.1f}%")
    print(f"{'='*50}")

    if errors:
        print(f"\nVehicles not found ({len(errors)}):")
        for e in errors:
            print(f"  - {e}")

    print(f"\nOutput saved to: {output_file}")


if __name__ == "__main__":
    main()
