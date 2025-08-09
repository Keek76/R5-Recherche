import os
import re
import pandas as pd
from serpapi import GoogleSearch
from email_service import send_email

SERPAPI_KEY = os.getenv("SERPAPI_KEY")

def search_leasing(query: str):
    params = {
        "q": query,
        "location": "Germany",
        "hl": "de",
        "gl": "de",
        "api_key": SERPAPI_KEY
    }
    search = GoogleSearch(params)
    results = search.get_dict()
    return results.get("organic_results", [])

def filter_offers(offers):
    filtered = []
    for offer in offers:
        text = (offer.get("title", "") + " " + offer.get("snippet", "")).lower()
        if "privat" in text and "€" in text:
            match = re.search(r"(\d+)[,.]?\d*\s*€", text)
            if match and float(match.group(1)) < 150:
                filtered.append(offer)
    return filtered

def save_to_excel(data, filename):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)

if __name__ == "__main__":
    raw_results = search_leasing("Renault 5 E-Tech Leasing Privat")
    filtered = filter_offers(raw_results)

    save_to_excel(filtered, "filtered_offers.xlsx")
    save_to_excel(raw_results, "all_offers.xlsx")

    if filtered:
        send_email("Gefundene Leasing-Angebote", "Siehe Anhang.", ["filtered_offers.xlsx"])
