import re
import requests
import pandas as pd

# =====================================================

# LOAD CATEGORY URLS

# =====================================================

categories_df = pd.read_csv("categories.csv")

category_urls = categories_df["url"].dropna().tolist()

# =====================================================

# VIDA XL WEBSHOPS

# =====================================================

shops = [
("GB","en","https://www.vidaxl.co.uk"),
("AE","ar","https://ar.vidaxl.ae"),
("AE","en","https://en.vidaxl.ae"),
("AU","en","https://www.vidaxl.com.au"),
("CA","en","https://en.vidaxl.ca"),
("IE","en","https://www.vidaxl.ie"),
("US","en","https://www.vidaxl.com"),
("AT","de","https://www.vidaxl.at"),
("CH","de","https://de.vidaxl.ch"),
("DE","de","https://www.vidaxl.de"),
("BE","fr","https://fr.vidaxl.be"),
("CA","fr","https://fr.vidaxl.ca"),
("CH","fr","https://fr.vidaxl.ch"),
("FR","fr","https://www.vidaxl.fr"),
("BE","nl","https://nl.vidaxl.be"),
("NL","nl","https://www.vidaxl.nl"),
("BG","bg","https://www.vidaxl.bg"),
("CZ","cs","https://www.vidaxl.cz"),
("DK","da","https://www.vidaxl.dk"),
("EE","et","https://www.vidaxl.ee"),
("ES","es","https://www.vidaxl.es"),
("FI","fi","https://www.vidaxl.fi"),
("GR","el","https://www.vidaxl.gr"),
("HR","hr","https://www.vidaxl.hr"),
("HU","hu","https://www.vidaxl.hu"),
("IS","is","https://is.vidaxl.is"),
("IT","it","https://www.vidaxl.it"),
("JP","ja","https://www.vidaxl.jp"),
("LT","lt","https://www.vidaxl.lt"),
("LV","lv","https://www.vidaxl.lv"),
("PL","pl","https://www.vidaxl.pl"),
("PT","pt","https://www.vidaxl.pt"),
("RO","ro","https://www.vidaxl.ro"),
("SE","sv","https://www.vidaxl.se"),
("SI","sl","https://www.vidaxl.si"),
("SK","sk","https://www.vidaxl.sk"),
("NO","nb","https://www.vidaxl.no")
]

# =====================================================

# EXTRACT CATEGORY IDS

# =====================================================

category_ids = []

for url in category_urls:
match = re.search(r'/g/(\d+)', url)

```
if match:
    category_ids.append(match.group(1))
```

print(f"Found {len(category_ids)} category IDs")

# =====================================================

# SCRAPE

# =====================================================

results = []

headers = {
"User-Agent": "Mozilla/5.0"
}

for category_id in category_ids:

```
print("=" * 80)
print(f"Processing Category ID: {category_id}")
print("=" * 80)

for country, language, base_url in shops:

    test_url = f"{base_url}/g/{category_id}"

    try:

        response = requests.get(
            test_url,
            headers=headers,
            allow_redirects=True,
            timeout=20
        )

        final_url = response.url

        print(f"{country}-{language}: {final_url}")

        results.append({
            "CategoryID": category_id,
            "Country": country,
            "Language": language,
            "BaseURL": base_url,
            "FinalURL": final_url,
            "StatusCode": response.status_code
        })

    except Exception as e:

        print(f"{country}-{language}: ERROR")

        results.append({
            "CategoryID": category_id,
            "Country": country,
            "Language": language,
            "BaseURL": base_url,
            "FinalURL": "",
            "StatusCode": "",
            "Error": str(e)
        })
```

# =====================================================

# EXPORT TO EXCEL

# =====================================================

df = pd.DataFrame(results)

output_file = "vidaxl_category_urls.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
df.to_excel(writer, sheet_name="Category URLs", index=False)

print(f"Excel created: {output_file}")

