ChatGPT Plus





Yes. Correct order:

Extract candidates → check TVA candidates → merge/clean → export.

Use this structure:

1. First install needed library
!pip install requests beautifulsoup4 openpyxl
2. Add this TVA candidate extractor
import re
import pandas as pd
import requests
from bs4 import BeautifulSoup
from pathlib import Path
from datetime import datetime
import time

def normalize_possible_tva(x):
    if pd.isna(x):
        return None

    s = str(x).upper().strip()
    s = re.sub(r"[^A-Z0-9]", "", s)

    if re.match(r"^BE[0-9]{9,10}$", s):
        return s

    if s.isdigit() and len(s) in [9, 10]:
        return "BE" + s.zfill(10)

    return None
3. Check TVA on KBO website
KBO Public Search shows public data for registered Belgian entities, so it is useful for validation. 

def check_tva_kbo(tva):
    if not tva:
        return False

    number = tva.replace("BE", "")

    url = "https://kbopub.economie.fgov.be/kbopub/zoeknummerform.html"

    try:
        r = requests.get(
            url,
            params={"nummer": number},
            timeout=10
        )

        text = r.text.lower()

        if "geen resultaat" in text or "aucun résultat" in text or "no result" in text:
            return False

        if number in text:
            return True

        return False

    except Exception:
        return False
4. Find TVA anywhere in the row
def find_valid_tva_in_row(row):
    candidates = []

    for value in row:
        candidate = normalize_possible_tva(value)
        if candidate:
            candidates.append(candidate)

    candidates = list(set(candidates))

    for candidate in candidates:
        if check_tva_kbo(candidate):
            return candidate

        time.sleep(0.5)  # be gentle with the website

    return None
5. Apply AFTER merging all files
Put this after:

final_df = pd.concat(all_data, ignore_index=True)
Then run:

final_df["TVA_CHECKED"] = final_df.apply(find_valid_tva_in_row, axis=1)

final_df["TVA"] = final_df["TVA_CHECKED"]

final_df.drop(columns=["TVA_CHECKED"], inplace=True)
6. Then clean phones and deduplicate
final_df["GSM"] = final_df["GSM"].apply(clean_phone)
final_df["TEL"] = final_df["TEL"].apply(clean_phone)

with_tva = final_df[final_df["TVA"].notna()].copy()
without_tva = final_df[final_df["TVA"].isna()].copy()

with_tva = with_tva.drop_duplicates(subset=["TVA"], keep="first")

without_tva = without_tva.drop_duplicates(
    subset=["STE", "CP", "CITY"],
    keep="first"
)

final_df = pd.concat([with_tva, without_tva], ignore_index=True)
7. Export
output_file = "master_companies_checked_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"

final_df.to_excel(output_file, index=False, engine="openpyxl")

print("DONE:", output_file)
Important: don’t check all rows if you have thousands. Start with a small test:

test_df = final_df.head(20).copy()
test_df["TVA_CHECKED"] = test_df.apply(find_valid_tva_in_row, axis=1)
test_df[["STE", "TVA_CHECKED", "GSM", "TEL"]]

yea  give me everything step by step since u know  the whole problem that i cite to u 

We’ll build it in a safer order: first merge all files into one raw master table, then collect possible TVA numbers from TVA/GSM/TEL/row values, then check suspicious candidates, then clean phones and deduplicate.

Yes. Use this full step-by-step code in a new notebook.

KBO Public Search is useful because it shows public data for entities registered in the Belgian company register. 

Step 1 — Install libraries
!pip install pandas openpyxl requests beautifulsoup4
Step 2 — Imports
import pandas as pd
import re
import time
import requests
from bs4 import BeautifulSoup
from pathlib import Path
from datetime import datetime
Step 3 — Load all Excel files
files = (
    list(Path(r"C:\Users\Sunlight\Desktop\DOSSIE 09-06-2026 - Copie\3-AB-12-08-2024-TR").rglob("*.xlsx"))
    +
    list(Path(r"C:\Users\Sunlight\Desktop\DOSSIE 09-06-2026 - Copie\Chaima 10-06-2024--TR").rglob("*.xlsx"))
    +
    list(Path(r"C:\Users\Sunlight\Desktop\DOSSIE 09-06-2026 - Copie\Imane 10-06-2024--TR").rglob("*.xlsx"))
)

print("Total files:", len(files))
Step 4 — Final columns
COLUMNS = [
    "STE",
    "NAME",
    "TVA",
    "ADDRESS",
    "HOUSE_NUM",
    "CP",
    "CITY",
    "GSM",
    "TEL",
    "FORME_JURIDIQUE",
    "CODE_LINGUISTIQUE",
    "CIVILITE",
    "COMMENTS",
    "OTHER_INFO",
    "SOURCE_FILE"
]
Step 5 — Column mapping
rename_map = {
    # STE
    "DENOMINATION SOCIALE": "STE",
    "DÉNOMINATION SOCIALE": "STE",
    "NAME_ORG1": "STE",
    "NAME_ORG2": "STE",
    "RS": "STE",
    "ENTITY NAME": "STE",
    "TITLE": "STE",

    # NAME
    "NOM": "NAME",
    "PRENOM": "NAME",
    "PRÉNOM": "NAME",
    "NAME": "NAME",

    # ADDRESS
    "RUE": "ADDRESS",
    "ADRESSE": "ADDRESS",
    "STREET60": "ADDRESS",
    "ADRESSE COMPLET": "ADDRESS",
    "ADRES": "ADDRESS",
    "STRAAT": "ADDRESS",

    # HOUSE NUMBER
    "NUMERO DE MAISON": "HOUSE_NUM",
    "NUMÉRO DE MAISON": "HOUSE_NUM",
    "AJOUT DE NUMERO DE MAISON": "HOUSE_NUM",
    "AJOUT DE NUMÉRO DE MAISON": "HOUSE_NUM",
    "NUMERO": "HOUSE_NUM",
    "NUMÉRO": "HOUSE_NUM",
    "HOUSE_NUM1": "HOUSE_NUM",
    "HOUSE_NUM2": "HOUSE_NUM",
    "NUM": "HOUSE_NUM",
    "N": "HOUSE_NUM",

    # CP
    "CODE POSTAL": "CP",
    "POSTALCODE": "CP",
    "ZIPCODE": "CP",
    "PC": "CP",
    "POST CODE": "CP",

    # CITY
    "LIEU(COMPLET)": "CITY",
    "PROVINCE": "CITY",
    "CITY-1": "CITY",
    "CITY": "CITY",
    "VILLE": "CITY",

    # TVA
    "NUMERO D'ENTREPRISE": "TVA",
    "NUMÉRO D'ENTREPRISE": "TVA",
    "VAT_REG_NO": "TVA",
    "TVA": "TVA",
    "FINALVAT": "TVA",
    "COLONNE6": "TVA",

    # GSM
    "GSM": "GSM",
    "PHONE_CLEAN": "GSM",
    "TEL2": "GSM",

    # TEL
    "NUMERO DE TELEPHONE": "TEL",
    "NUMÉRO DE TÉLÉPHONE": "TEL",
    "TELEPHONE": "TEL",
    "TÉLÉPHONE": "TEL",
    "PHONE": "TEL",
    "FIX": "TEL",
    "NO DE TELEPHONE": "TEL",
    "NO DE TÉLÉPHONE": "TEL",
    "FINAL PHONE": "TEL",
    "COLONNE5": "TEL",
    "COLONNE7": "TEL",

    # COMMENTS
    "DESCRIPTION": "COMMENTS"
}
Step 6 — Helper functions
def clean_text(x):
    if pd.isna(x):
        return None
    x = str(x).strip()
    if x.lower() in ["nan", "none", ""]:
        return None
    return x


def normalize_tva_candidate(x):
    """
    Converts possible TVA:
    BE0462890433 -> BE0462890433
    0462890433   -> BE0462890433
    462890433    -> BE0462890433
    """
    if pd.isna(x):
        return None

    s = str(x).upper().strip()
    s = re.sub(r"[^A-Z0-9]", "", s)

    if s == "":
        return None

    if re.match(r"^BE[0-9]{9,10}$", s):
        return s

    if s.isdigit() and len(s) in [9, 10]:
        return "BE" + s.zfill(10)

    return None


def clean_phone(x):
    """
    Cleans GSM/TEL.
    Keeps Belgian phone-like numbers.
    Removes long EAN/barcode noise.
    """
    if pd.isna(x):
        return None

    s = str(x).strip()
    if s.lower() in ["nan", "none", ""]:
        return None

    digits = re.sub(r"\D", "", s)

    if digits == "":
        return None

    # Remove EAN/noise
    if len(digits) > 12:
        return None

    # 0032xxxx -> 32xxxx
    if digits.startswith("0032"):
        digits = "32" + digits[4:]

    # 0xxxx -> 32xxxx
    elif digits.startswith("0"):
        digits = "32" + digits[1:]

    # Belgian mobile/fixed after cleaning
    if digits.startswith("32") and len(digits) in [10, 11]:
        return digits

    # Local fixed number, like 22191358
    if len(digits) in [8, 9]:
        return digits

    return None


def is_ean_noise(x):
    if pd.isna(x):
        return False

    digits = re.sub(r"\D", "", str(x))

    return len(digits) >= 13
Step 7 — KBO check function
def check_tva_kbo(tva):
    """
    Checks a Belgian company number on KBO Public Search.
    Returns True if the number appears valid/found, False otherwise.

    Important: use this only for suspicious candidates, not millions of rows.
    """
    if not tva:
        return False

    number = str(tva).upper().replace("BE", "")
    number = re.sub(r"\D", "", number)

    if len(number) not in [9, 10]:
        return False

    url = "https://kbopub.economie.fgov.be/kbopub/zoeknummerform.html"

    try:
        response = requests.get(
            url,
            params={"nummer": number},
            timeout=15,
            headers={
                "User-Agent": "Mozilla/5.0"
            }
        )

        text = response.text.lower()

        # If the number appears in returned page, likely found
        if number in re.sub(r"\D", "", text):
            return True

        invalid_words = [
            "geen resultaat",
            "aucun résultat",
            "no result",
            "niet gevonden",
            "introuvable"
        ]

        if any(word in text for word in invalid_words):
            return False

        return False

    except Exception:
        return False
Step 8 — Find TVA from the whole row
This solves the problem where TVA is accidentally in GSM or TEL.

def find_tva_candidates_in_row(row):
    candidates = []

    for value in row.values:
        candidate = normalize_tva_candidate(value)
        if candidate:
            candidates.append(candidate)

    return list(dict.fromkeys(candidates))


def choose_best_tva(row):
    """
    Priority:
    1. TVA column if valid and KBO confirms
    2. Any candidate in the row if KBO confirms
    3. Return None if no confirmed TVA
    """

    candidates = []

    # Priority: current TVA column first
    if "TVA" in row.index:
        first = normalize_tva_candidate(row["TVA"])
        if first:
            candidates.append(first)

    # Then scan whole row
    for candidate in find_tva_candidates_in_row(row):
        if candidate not in candidates:
            candidates.append(candidate)

    for candidate in candidates:
        if check_tva_kbo(candidate):
            time.sleep(0.4)
            return candidate

        time.sleep(0.4)

    return None
Step 9 — Classify file type
def classify_file(file):
    try:
        df = pd.read_excel(file, nrows=5, engine="openpyxl")
        cols = [str(c).strip().upper() for c in df.columns]

        if any(col in rename_map for col in cols):
            return "TYPE_1_HEADER"

        return "TYPE_2_NO_HEADER"

    except:
        return "TYPE_2_NO_HEADER"
Step 10 — Process files with headers
def process_type_1(file):
    df = pd.read_excel(file, engine="openpyxl")

    df.columns = [str(c).strip().upper() for c in df.columns]
    df.rename(columns=rename_map, inplace=True)

    # Merge duplicated columns after renaming
    df = df.groupby(level=0, axis=1).first()

    other_cols = [c for c in df.columns if c not in COLUMNS]

    if other_cols:
        df["OTHER_INFO"] = df[other_cols].apply(
            lambda row: " | ".join(
                f"{col}: {val}" for col, val in row.items()
                if pd.notna(val) and str(val).strip().lower() not in ["nan", ""]
            ),
            axis=1
        )
    else:
        df["OTHER_INFO"] = None

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None

    df["SOURCE_FILE"] = file.name

    return df[COLUMNS]
Step 11 — Process files without headers
For your main format:

STE | ADDRESS | CP | CITY | TEL | TVA

def process_type_2(file):
    df = pd.read_excel(file, header=None, engine="openpyxl")
    df = df.dropna(how="all")

    df = df.iloc[:, :6]

    while df.shape[1] < 6:
        df[df.shape[1]] = None

    df.columns = ["STE", "ADDRESS", "CP", "CITY", "TEL", "TVA"]

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None

    df["SOURCE_FILE"] = file.name

    return df[COLUMNS]
Step 12 — Process all files
all_data = []
errors = []

for i, file in enumerate(files, start=1):
    try:
        file_type = classify_file(file)

        if file_type == "TYPE_1_HEADER":
            df = process_type_1(file)
        else:
            df = process_type_2(file)

        all_data.append(df)

        print(f"{i}/{len(files)} processed - {file_type} - {file.name}")

    except Exception as e:
        errors.append((file.name, str(e)))
        print("ERROR:", file.name, "->", e)

print("Processed files:", len(all_data))
print("Errors:", len(errors))
Step 13 — Merge all files
final_df = pd.concat(all_data, ignore_index=True)

print("Rows before cleaning:", len(final_df))
Step 14 — Basic cleaning
for col in ["STE", "NAME", "ADDRESS", "HOUSE_NUM", "CP", "CITY",
            "FORME_JURIDIQUE", "CODE_LINGUISTIQUE", "CIVILITE",
            "COMMENTS", "OTHER_INFO"]:
    final_df[col] = final_df[col].apply(clean_text)
Step 15 — Remove fake header rows
bad_words = (
    "numero|numéro|denomination|dénomination|telephone|téléphone|"
    "adresse|postal|vat|tva|société|societe"
)

final_df = final_df[
    ~final_df["STE"].astype(str).str.lower().str.contains(bad_words, na=False)
]

print("Rows after removing fake headers:", len(final_df))
Step 16 — First clean phones
final_df["GSM"] = final_df["GSM"].apply(clean_phone)
final_df["TEL"] = final_df["TEL"].apply(clean_phone)
Step 17 — Create TVA candidates before online checking
This exports a small file so you can see what Python wants to check.

final_df["TVA_CANDIDATES"] = final_df.apply(
    lambda row: ", ".join(find_tva_candidates_in_row(row)),
    axis=1
)

to_check = final_df[final_df["TVA_CANDIDATES"] != ""].copy()

to_check[["STE", "TVA", "GSM", "TEL", "TVA_CANDIDATES", "SOURCE_FILE"]].to_excel(
    "tva_candidates_before_kbo_check.xlsx",
    index=False,
    engine="openpyxl"
)

print("Candidates exported:", len(to_check))
Step 18 — Test KBO check on only 10 rows first
Do this before running on all data.

test_df = final_df.head(10).copy()

test_df["TVA_CHECKED"] = test_df.apply(choose_best_tva, axis=1)

test_df[["STE", "TVA", "GSM", "TEL", "TVA_CANDIDATES", "TVA_CHECKED"]]
If this works, continue.

Step 19 — Check TVA using KBO
final_df["TVA_CHECKED"] = final_df.apply(choose_best_tva, axis=1)

final_df["TVA"] = final_df["TVA_CHECKED"]

final_df.drop(columns=["TVA_CHECKED"], inplace=True)

print("Valid checked TVA:", final_df["TVA"].notna().sum())
Step 20 — Remove EAN noise
final_df["OTHER_INFO"] = final_df["OTHER_INFO"].apply(
    lambda x: None if is_ean_noise(x) else x
)

final_df = final_df[
    final_df["STE"].notna() |
    final_df["TVA"].notna() |
    final_df["GSM"].notna() |
    final_df["TEL"].notna()
]

print("Rows after EAN/noise cleaning:", len(final_df))
Step 21 — Deduplicate by TVA
with_tva = final_df[final_df["TVA"].notna()].copy()
without_tva = final_df[final_df["TVA"].isna()].copy()

with_tva = with_tva.drop_duplicates(subset=["TVA"], keep="first")

without_tva = without_tva.drop_duplicates(
    subset=["STE", "CP", "CITY"],
    keep="first"
)

final_df = pd.concat([with_tva, without_tva], ignore_index=True)

print("Rows after deduplication:", len(final_df))
Step 22 — Export final file
output_file = "master_companies_checked_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"

final_df.to_excel(
    output_file,
    index=False,
    engine="openpyxl"
)

print("DONE:", output_file)
Step 23 — Check result
final_df[["STE", "TVA", "GSM", "TEL", "SOURCE_FILE"]].head(30)
Important: the KBO checking step can be slow. If you have thousands of rows, start with head(100) first, then run all rows after you confirm it works.


rows before  cleaning 889620
rows after text cleaning 889620
columns 15 
rows after removing  fake headers 886102
candidates exported 771816 

That is too many rows to check online one by one.

You have:

Rows: 886,102
TVA candidates: 771,816
If Python checks KBO for 771k candidates, it will take days and the website may block you.

Do this instead: check only UNIQUE candidates
Run this:

unique_candidates = (
    final_df["TVA_CANDIDATES"]
    .dropna()
    .astype(str)
    .str.split(", ")
    .explode()
    .dropna()
    .drop_duplicates()
)

unique_candidates = unique_candidates[unique_candidates != ""]

print("Unique TVA candidates:", len(unique_candidates))
Then check only a small sample first:

sample_candidates = unique_candidates.head(100)

checked_results = []

for i, tva in enumerate(sample_candidates, start=1):
    valid = check_tva_kbo(tva)
    checked_results.append((tva, valid))
    print(i, tva, valid)
    time.sleep(0.5)

checked_df = pd.DataFrame(checked_results, columns=["TVA", "IS_VALID_KBO"])
checked_df
If this works, then check all unique candidates, not all rows:

checked_results = []

for i, tva in enumerate(unique_candidates, start=1):
    valid = check_tva_kbo(tva)
    checked_results.append((tva, valid))

    if i % 100 == 0:
        print(i, "/", len(unique_candidates))

    time.sleep(0.5)

checked_df = pd.DataFrame(checked_results, columns=["TVA", "IS_VALID_KBO"])
Then merge result back:

valid_tvas = set(checked_df.loc[checked_df["IS_VALID_KBO"], "TVA"])

def choose_checked_tva_from_candidates(candidates):
    if pd.isna(candidates) or candidates == "":
        return None

    for tva in str(candidates).split(", "):
        if tva in valid_tvas:
            return tva

    return None

final_df["TVA"] = final_df["TVA_CANDIDATES"].apply(choose_checked_tva_from_candidates)
But first run:

print("Unique TVA candidates:", len(unique_candidates))
If it is still more than 50,000, we should not use the website directly. We should reduce candidates by checking only:

candidates from original TVA column

candidates from TEL/GSM only when TVA is empty

candidates from rows where STE is not empty.


i apply unique condidates after i run tva condidats 

Yes 👍

The correct order is:

1️⃣ First run
final_df["TVA_CANDIDATES"] = final_df.apply(
    lambda row: ", ".join(find_tva_candidates_in_row(row)),
    axis=1
)
This creates the column:

TVA_CANDIDATES
2️⃣ Then run
unique_candidates = (
    final_df["TVA_CANDIDATES"]
    .dropna()
    .astype(str)
    .str.split(", ")
    .explode()
    .dropna()
    .drop_duplicates()
)

unique_candidates = unique_candidates[
    unique_candidates != ""
]

print("Unique TVA candidates:", len(unique_candidates))
3️⃣ Then tell me the result of:
print("Unique TVA candidates:", len(unique_candidates))
This number is extremely important.

Because:

If it is 500 → we can check all of them.

If it is 5,000 → maybe.

If it is 50,000+ → we need another strategy.

If it is 100,000+ → absolutely do not query the website.

Before going further
I have another concern.

You have:

886,102 rows
from only:

182 Excel files
That is huge.

It makes me think many Type 2 files contain:

repeated rows

blank rows

address-only rows

imported database exports

Before we start validating TVA online, run:

print(final_df[["STE","TVA","GSM","TEL"]].head(20))
and show me a screenshot.

I want to see if the extraction is correct first.

Because if the extraction is wrong, validating TVA now will just validate bad data and waste hours.


they give me 689436

689,436 is far too many to check with the website. Don’t run KBO validation on that list.

It means our candidate search is too broad and is catching phone numbers/EAN/noise.

Do this instead: restrict TVA candidates
Replace your normalize_tva_candidate with this stricter version:

def normalize_tva_candidate(x):
    if pd.isna(x):
        return None

    s = str(x).upper().strip()
    s = re.sub(r"[^A-Z0-9]", "", s)

    if s == "":
        return None

    # Strong TVA: BE + 10 digits
    if re.match(r"^BE[0-9]{10}$", s):
        return s

    # Numeric TVA only if exactly 10 digits and starts with 0
    # Example: 0462890433 -> BE0462890433
    if s.isdigit() and len(s) == 10 and s.startswith("0"):
        return "BE" + s

    # 9-digit TVA only if it is NOT likely GSM
    # Example: 462890433 -> BE0462890433
    # But 478123456 could be GSM, so exclude common GSM prefixes
    if s.isdigit() and len(s) == 9:
        if s.startswith(("46", "47", "48", "49")):
            return None
        return "BE0" + s

    return None
Then recreate candidates:

final_df["TVA_CANDIDATES"] = final_df.apply(
    lambda row: ", ".join(find_tva_candidates_in_row(row)),
    axis=1
)

unique_candidates = (
    final_df["TVA_CANDIDATES"]
    .dropna()
    .astype(str)
    .str.split(", ")
    .explode()
    .dropna()
    .drop_duplicates()
)

unique_candidates = unique_candidates[unique_candidates != ""]

print("Unique TVA candidates:", len(unique_candidates))
If the number is still huge, we should check only candidates from the TVA column first, not from the whole row.


what about see how many one doesnt start with the two letter and we checked just those 

Yes, exactly. The best logic is:

Check existing TVA first

If TVA is numeric, try it with KBO before adding BE

If valid → save as BE + number

If not valid → scan the whole row for another TVA candidate

If no valid TVA found → keep TVA as None

Use this replacement function:

def validate_tva_with_kbo_or_row(row):
    # 1) Try TVA column first
    main_tva = normalize_tva_candidate(row.get("TVA"))

    if main_tva and check_tva_kbo(main_tva):
        return main_tva

    # 2) If TVA column failed, search all row values
    candidates = find_tva_candidates_in_row(row)

    for candidate in candidates:
        if check_tva_kbo(candidate):
            return candidate

        time.sleep(0.4)

    # 3) Nothing valid found
    return None
Then in the cleaning phase, replace this:

final_df["TVA"] = final_df["TVA"].apply(clean_tva)
with this:

final_df["TVA"] = final_df.apply(validate_tva_with_kbo_or_row, axis=1)
But important: do not apply it directly on 886,000 rows. It will be extremely slow.

Use it only on rows where TVA is missing or suspicious:

def is_tva_suspicious(x):
    if pd.isna(x):
        return True

    s = str(x).upper().strip()
    s = re.sub(r"[^A-Z0-9]", "", s)

    # Already clean BE + 10 digits
    if re.match(r"^BE[0-9]{10}$", s):
        return False

    return True
Then:

mask = final_df["TVA"].apply(is_tva_suspicious)

print("Rows needing KBO check:", mask.sum())

final_df.loc[mask, "TVA"] = final_df.loc[mask].apply(
    validate_tva_with_kbo_or_row,
    axis=1
)
This is the safest idea:

If TVA is already BE + 10 digits → keep it
If TVA is suspicious/null → check KBO
If KBO finds a valid number somewhere in the row → use it
If not → TVA = None
But first run:

print("Rows needing KBO check:", mask.sum())
If that number is still huge, we need to reduce it further before using the website.


i will do this after rename map step riht 

No ❌

Not after rename_map.

The order should be:

1. Imports
2. File paths
3. rename_map
4. Helper functions
clean_phone()

normalize_tva_candidate()

check_tva_kbo()

find_tva_candidates_in_row()

validate_tva_with_kbo_or_row()

5. Process files
process_type_1()

process_type_2()

6. Merge
final_df = pd.concat(all_data, ignore_index=True)
7. Basic cleaning
clean_text()
remove fake headers
clean phones
8. THEN create the suspicious mask
mask = final_df["TVA"].apply(is_tva_suspicious)

print("Rows needing KBO check:", mask.sum())
9. THEN apply KBO validation
final_df.loc[mask, "TVA"] = final_df.loc[mask].apply(
    validate_tva_with_kbo_or_row,
    axis=1
)
10. Deduplicate
drop_duplicates(subset=["TVA"])
11. Export
Why?

Because if you do KBO checking immediately after rename_map, the data is not even loaded yet. rename_map is just a dictionary.

You need:

Read files
→ Create final_df
→ Clean data
→ KBO validation
→ Deduplicate
→ Export
Before you run KBO validation, I want you to do one thing:

mask = final_df["TVA"].apply(is_tva_suspicious)

print("Rows needing KBO check:", mask.sum())
Tell me that number first.

If it says:

Rows needing KBO check: 5,000
great.

If it says:

Rows needing KBO check: 500,000
then we need another strategy because KBO validation will take forever.


its confusing could u give me the whole code organized lets seperate be first and check the left tva with kbo and see

Yes. Use this clean organized version. The logic is:

Merge all Excel files

Keep TVA that already starts with BE

For the rest, create possible TVA candidates

Check only those with KBO

Clean GSM/TEL

Deduplicate

Export

# =========================
# 1) INSTALL
# =========================
!pip install pandas openpyxl requests beautifulsoup4
# =========================
# 2) IMPORTS
# =========================
import pandas as pd
import re
import time
import requests
from pathlib import Path
from datetime import datetime
# =========================
# 3) FILE PATHS
# =========================
files = (
    list(Path(r"C:\Users\Sunlight\Desktop\DOSSIE 09-06-2026 - Copie\3-AB-12-08-2024-TR").rglob("*.xlsx"))
    +
    list(Path(r"C:\Users\Sunlight\Desktop\DOSSIE 09-06-2026 - Copie\Chaima 10-06-2024--TR").rglob("*.xlsx"))
    +
    list(Path(r"C:\Users\Sunlight\Desktop\DOSSIE 09-06-2026 - Copie\Imane 10-06-2024--TR").rglob("*.xlsx"))
)

print("Total files:", len(files))
# =========================
# 4) FINAL COLUMNS
# =========================
COLUMNS = [
    "STE", "NAME", "TVA", "ADDRESS", "HOUSE_NUM", "CP", "CITY",
    "GSM", "TEL", "FORME_JURIDIQUE", "CODE_LINGUISTIQUE",
    "CIVILITE", "COMMENTS", "OTHER_INFO", "SOURCE_FILE"
]
# =========================
# 5) RENAME MAP
# =========================
rename_map = {
    "DENOMINATION SOCIALE": "STE",
    "DÉNOMINATION SOCIALE": "STE",
    "NAME_ORG1": "STE",
    "NAME_ORG2": "STE",
    "RS": "STE",
    "ENTITY NAME": "STE",
    "TITLE": "STE",

    "NOM": "NAME",
    "PRENOM": "NAME",
    "PRÉNOM": "NAME",
    "NAME": "NAME",

    "RUE": "ADDRESS",
    "ADRESSE": "ADDRESS",
    "STREET60": "ADDRESS",
    "ADRESSE COMPLET": "ADDRESS",
    "ADRES": "ADDRESS",
    "STRAAT": "ADDRESS",

    "NUMERO DE MAISON": "HOUSE_NUM",
    "NUMÉRO DE MAISON": "HOUSE_NUM",
    "AJOUT DE NUMERO DE MAISON": "HOUSE_NUM",
    "AJOUT DE NUMÉRO DE MAISON": "HOUSE_NUM",
    "NUMERO": "HOUSE_NUM",
    "NUMÉRO": "HOUSE_NUM",
    "HOUSE_NUM1": "HOUSE_NUM",
    "HOUSE_NUM2": "HOUSE_NUM",
    "NUM": "HOUSE_NUM",
    "N": "HOUSE_NUM",

    "CODE POSTAL": "CP",
    "POSTALCODE": "CP",
    "ZIPCODE": "CP",
    "PC": "CP",
    "POST CODE": "CP",

    "LIEU(COMPLET)": "CITY",
    "PROVINCE": "CITY",
    "CITY-1": "CITY",
    "CITY": "CITY",
    "VILLE": "CITY",

    "NUMERO D'ENTREPRISE": "TVA",
    "NUMÉRO D'ENTREPRISE": "TVA",
    "VAT_REG_NO": "TVA",
    "TVA": "TVA",
    "FINALVAT": "TVA",

    "GSM": "GSM",
    "PHONE_CLEAN": "GSM",
    "TEL2": "GSM",

    "NUMERO DE TELEPHONE": "TEL",
    "NUMÉRO DE TÉLÉPHONE": "TEL",
    "TELEPHONE": "TEL",
    "TÉLÉPHONE": "TEL",
    "PHONE": "TEL",
    "FIX": "TEL",
    "NO DE TELEPHONE": "TEL",
    "NO DE TÉLÉPHONE": "TEL",
    "FINAL PHONE": "TEL",

    "DESCRIPTION": "COMMENTS"
}
# =========================
# 6) HELPER FUNCTIONS
# =========================
def clean_text(x):
    if pd.isna(x):
        return None
    x = str(x).strip()
    if x.lower() in ["nan", "none", ""]:
        return None
    return x


def clean_phone(x):
    if pd.isna(x):
        return None

    s = str(x).strip()
    if s.lower() in ["nan", "none", ""]:
        return None

    digits = re.sub(r"\D", "", s)

    if digits == "":
        return None

    # remove EAN/barcode-like noise
    if len(digits) > 12:
        return None

    # 0032xxxx -> 32xxxx
    if digits.startswith("0032"):
        digits = "32" + digits[4:]

    # 0xxxx -> 32xxxx
    elif digits.startswith("0"):
        digits = "32" + digits[1:]

    if digits.startswith("32") and len(digits) in [10, 11]:
        return digits

    if len(digits) in [8, 9]:
        return digits

    return None


def normalize_be_tva(x):
    """
    Only keep TVA already starting with BE.
    """
    if pd.isna(x):
        return None

    s = str(x).upper().strip()
    s = re.sub(r"[^A-Z0-9]", "", s)

    if re.match(r"^BE[0-9]{9,10}$", s):
        return "BE" + re.sub(r"\D", "", s).zfill(10)

    return None


def normalize_numeric_tva_candidate(x):
    """
    Prepare possible numeric TVA for KBO check.
    Does NOT automatically trust it.
    """
    if pd.isna(x):
        return None

    s = str(x).upper().strip()
    s = re.sub(r"[^0-9]", "", s)

    if len(s) == 10:
        return "BE" + s

    if len(s) == 9:
        return "BE" + s.zfill(10)

    return None


def is_ean_noise(x):
    if pd.isna(x):
        return False

    digits = re.sub(r"\D", "", str(x))
    return len(digits) >= 13
# =========================
# 7) KBO CHECK
# =========================
def check_tva_kbo(tva):
    if not tva:
        return False

    number = str(tva).upper().replace("BE", "")
    number = re.sub(r"\D", "", number)

    if len(number) not in [9, 10]:
        return False

    url = "https://kbopub.economie.fgov.be/kbopub/zoeknummerform.html"

    try:
        response = requests.get(
            url,
            params={"nummer": number},
            timeout=15,
            headers={"User-Agent": "Mozilla/5.0"}
        )

        text_digits = re.sub(r"\D", "", response.text)

        return number in text_digits

    except Exception:
        return False
# =========================
# 8) FILE CLASSIFICATION
# =========================
def classify_file(file):
    try:
        df = pd.read_excel(file, nrows=5, engine="openpyxl")
        cols = [str(c).strip().upper() for c in df.columns]

        if any(col in rename_map for col in cols):
            return "TYPE_1_HEADER"

        return "TYPE_2_NO_HEADER"

    except:
        return "TYPE_2_NO_HEADER"
# =========================
# 9) PROCESS TYPE 1: WITH HEADERS
# =========================
def process_type_1(file):
    df = pd.read_excel(file, engine="openpyxl")

    df.columns = [str(c).strip().upper() for c in df.columns]
    df.rename(columns=rename_map, inplace=True)

    df = df.groupby(level=0, axis=1).first()

    other_cols = [c for c in df.columns if c not in COLUMNS]

    if other_cols:
        df["OTHER_INFO"] = df[other_cols].apply(
            lambda row: " | ".join(
                f"{col}: {val}" for col, val in row.items()
                if pd.notna(val) and str(val).strip().lower() not in ["nan", ""]
            ),
            axis=1
        )
    else:
        df["OTHER_INFO"] = None

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None

    df["SOURCE_FILE"] = file.name

    return df[COLUMNS]
# =========================
# 10) PROCESS TYPE 2: WITHOUT HEADERS
# Format: STE | ADDRESS | CP | CITY | TEL | TVA
# =========================
def process_type_2(file):
    df = pd.read_excel(file, header=None, engine="openpyxl")
    df = df.dropna(how="all")

    df = df.iloc[:, :6]

    while df.shape[1] < 6:
        df[df.shape[1]] = None

    df.columns = ["STE", "ADDRESS", "CP", "CITY", "TEL", "TVA"]

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None

    df["SOURCE_FILE"] = file.name

    return df[COLUMNS]
# =========================
# 11) PROCESS ALL FILES
# =========================
all_data = []
errors = []

for i, file in enumerate(files, start=1):
    try:
        file_type = classify_file(file)

        if file_type == "TYPE_1_HEADER":
            df = process_type_1(file)
        else:
            df = process_type_2(file)

        all_data.append(df)
        print(f"{i}/{len(files)} processed - {file_type} - {file.name}")

    except Exception as e:
        errors.append((file.name, str(e)))
        print("ERROR:", file.name, "->", e)

print("Processed files:", len(all_data))
print("Errors:", len(errors))
# =========================
# 12) MERGE
# =========================
final_df = pd.concat(all_data, ignore_index=True)

print("Rows before cleaning:", len(final_df))
print("Columns:", len(final_df.columns))
# =========================
# 13) BASIC TEXT CLEANING
# =========================
for col in ["STE", "NAME", "ADDRESS", "HOUSE_NUM", "CP", "CITY",
            "FORME_JURIDIQUE", "CODE_LINGUISTIQUE", "CIVILITE",
            "COMMENTS", "OTHER_INFO"]:
    final_df[col] = final_df[col].apply(clean_text)

print("Rows after text cleaning:", len(final_df))
# =========================
# 14) REMOVE FAKE HEADER ROWS
# =========================
bad_words = (
    "numero|numéro|denomination|dénomination|telephone|téléphone|"
    "adresse|postal|vat|tva|société|societe"
)

final_df = final_df[
    ~final_df["STE"].astype(str).str.lower().str.contains(bad_words, na=False)
]

print("Rows after removing fake headers:", len(final_df))
# =========================
# 15) SEPARATE TVA THAT ALREADY STARTS WITH BE
# =========================
final_df["TVA_BE_CLEAN"] = final_df["TVA"].apply(normalize_be_tva)

already_be = final_df["TVA_BE_CLEAN"].notna().sum()
not_be = final_df["TVA_BE_CLEAN"].isna().sum()

print("TVA already starts with BE:", already_be)
print("TVA not BE / missing / suspicious:", not_be)
# =========================
# 16) CREATE CANDIDATES ONLY FOR NON-BE TVA ROWS
# Search TVA, GSM, TEL, OTHER_INFO only.
# =========================
def find_candidates_for_suspicious_row(row):
    candidates = []

    search_cols = ["TVA", "GSM", "TEL", "OTHER_INFO"]

    for col in search_cols:
        value = row.get(col)
        candidate = normalize_numeric_tva_candidate(value)

        if candidate:
            candidates.append(candidate)

    return list(dict.fromkeys(candidates))


suspicious_mask = final_df["TVA_BE_CLEAN"].isna()

final_df.loc[suspicious_mask, "TVA_CANDIDATES"] = final_df.loc[suspicious_mask].apply(
    lambda row: ", ".join(find_candidates_for_suspicious_row(row)),
    axis=1
)

final_df.loc[~suspicious_mask, "TVA_CANDIDATES"] = ""

print("Rows needing candidate search:", suspicious_mask.sum())
print("Rows with candidates:", (final_df["TVA_CANDIDATES"] != "").sum())
# =========================
# 17) UNIQUE CANDIDATES TO CHECK WITH KBO
# =========================
unique_candidates = (
    final_df.loc[suspicious_mask, "TVA_CANDIDATES"]
    .dropna()
    .astype(str)
    .str.split(", ")
    .explode()
    .str.strip()
    .drop_duplicates()
)

unique_candidates = unique_candidates[unique_candidates != ""]

print("Unique candidates to check with KBO:", len(unique_candidates))
print(unique_candidates.head(20))
# =========================
# 18) TEST KBO CHECK ON FIRST 20 CANDIDATES
# =========================
test_candidates = unique_candidates.head(20)

test_results = []

for i, tva in enumerate(test_candidates, start=1):
    valid = check_tva_kbo(tva)
    test_results.append((tva, valid))
    print(i, tva, valid)
    time.sleep(0.5)

test_results_df = pd.DataFrame(test_results, columns=["TVA_CANDIDATE", "IS_VALID_KBO"])
test_results_df
# =========================
# 19) CHECK ALL UNIQUE CANDIDATES WITH KBO
# Run this only if Step 18 works.
# =========================
checked_results = []

for i, tva in enumerate(unique_candidates, start=1):
    valid = check_tva_kbo(tva)
    checked_results.append((tva, valid))

    if i % 100 == 0:
        print(i, "/", len(unique_candidates))

    time.sleep(0.5)

checked_df = pd.DataFrame(
    checked_results,
    columns=["TVA_CANDIDATE", "IS_VALID_KBO"]
)

checked_df.to_excel("kbo_checked_tva_candidates.xlsx", index=False)

print("KBO checking done")
# =========================
# 20) APPLY KBO RESULTS BACK TO FINAL_DF
# =========================
valid_tvas = set(
    checked_df.loc[checked_df["IS_VALID_KBO"], "TVA_CANDIDATE"]
)

def choose_checked_tva(row):
    # Keep BE if already clean
    if pd.notna(row["TVA_BE_CLEAN"]):
        return row["TVA_BE_CLEAN"]

    candidates = str(row.get("TVA_CANDIDATES", "")).split(", ")

    for candidate in candidates:
        if candidate in valid_tvas:
            return candidate

    return None


final_df["TVA"] = final_df.apply(choose_checked_tva, axis=1)

print("Final valid TVA:", final_df["TVA"].notna().sum())
# =========================
# 21) CLEAN GSM/TEL AFTER TVA DECISION
# =========================
final_df["GSM"] = final_df["GSM"].apply(clean_phone)
final_df["TEL"] = final_df["TEL"].apply(clean_phone)
# =========================
# 22) REMOVE EAN NOISE
# =========================
final_df["OTHER_INFO"] = final_df["OTHER_INFO"].apply(
    lambda x: None if is_ean_noise(x) else x
)

final_df = final_df[
    final_df["STE"].notna() |
    final_df["TVA"].notna() |
    final_df["GSM"].notna() |
    final_df["TEL"].notna()
]

print("Rows after EAN/noise cleaning:", len(final_df))
# =========================
# 23) DEDUPLICATE BY TVA
# =========================
with_tva = final_df[final_df["TVA"].notna()].copy()
without_tva = final_df[final_df["TVA"].isna()].copy()

with_tva = with_tva.drop_duplicates(subset=["TVA"], keep="first")

without_tva = without_tva.drop_duplicates(
    subset=["STE", "CP", "CITY"],
    keep="first"
)

final_df = pd.concat([with_tva, without_tva], ignore_index=True)

print("Rows after deduplication:", len(final_df))
# =========================
# 24) DROP TEMP COLUMNS
# =========================
final_df = final_df.drop(
    columns=["TVA_BE_CLEAN", "TVA_CANDIDATES"],
    errors="ignore"
)
# =========================
# 25) EXPORT
# =========================
output_file = "master_companies_checked_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"

final_df.to_excel(
    output_file,
    index=False,
    engine="openpyxl"
)

print("DONE:", output_file)
# =========================
# 26) CHECK RESULT
# =========================
final_df[["STE", "TVA", "GSM", "TEL", "SOURCE_FILE"]].head(30)
Important: after Step 17, if Unique candidates to check with KBO is still huge, don’t run Step 19 yet.


it gives me 607002 its too much for kbo to check so the plan is treat each folder alone since we have three folders i am gonna treat each one alone and then iwill merge them manually using excel 

Yes, that’s a good practical plan.

Since 607,002 candidates is too much for KBO, treating each folder separately is smarter. But I suggest:

Don’t merge manually in Excel unless necessary.
Create 3 clean Excel outputs, then merge them with Excel or Python later.

Use this structure:

Folder 1 → master_AB_clean.xlsx
Folder 2 → master_Chaima_clean.xlsx
Folder 3 → master_Imane_clean.xlsx
Change only the file path part each time
For folder 1:

files = list(Path(r"C:\Users\Sunlight\Desktop\DOSSIE 09-06-2026 - Copie\3-AB-12-08-2024-TR").rglob("*.xlsx"))
print("Total files:", len(files))
For folder 2:

files = list(Path(r"C:\Users\Sunlight\Desktop\DOSSIE 09-06-2026 - Copie\Chaima 10-06-2024--TR").rglob("*.xlsx"))
print("Total files:", len(files))
For folder 3:

files = list(Path(r"C:\Users\Sunlight\Desktop\DOSSIE 09-06-2026 - Copie\Imane 10-06-2024--TR").rglob("*.xlsx"))
print("Total files:", len(files))
Then run the same cleaning pipeline for each folder and export with a different name:

output_file = "master_AB_clean.xlsx"
then:

output_file = "master_Chaima_clean.xlsx"
then:

output_file = "master_Imane_clean.xlsx"
After that, you can merge the 3 final files.

To merge the 3 final files in Python
files_final = [
    "master_AB_clean.xlsx",
    "master_Chaima_clean.xlsx",
    "master_Imane_clean.xlsx"
]

dfs = [pd.read_excel(f) for f in files_final]

final_all = pd.concat(dfs, ignore_index=True)

final_all = final_all.drop_duplicates(subset=["TVA"], keep="first")

final_all.to_excel("FINAL_ALL_COMPANIES.xlsx", index=False)

print("DONE")
This is better than manually merging in Excel because it avoids copy-paste mistakes.


wait but first give me the whole code for one folder with a stricted kbo because when i checked its false so could u make it more accurate now give the whole code and then i will apply the same of the two other folder 

I’ll give you a one-folder version and make KBO stricter by checking the returned page for the company number in multiple formats, not only raw digits. It will also save a separate file of candidates checked so you can audit what KBO accepted/refused.

Yes. Use this for one folder only first.

The KBO website is official and shows public data for entities registered in Belgium’s company register. 

# ==========================================================
# 1) INSTALL
# ==========================================================
!pip install pandas openpyxl requests beautifulsoup4
# ==========================================================
# 2) IMPORTS
# ==========================================================
import pandas as pd
import re
import time
import requests
from pathlib import Path
from datetime import datetime
from bs4 import BeautifulSoup
# ==========================================================
# 3) CHOOSE ONE FOLDER ONLY
# Change this path for folder 2 and folder 3 later
# ==========================================================
folder_path = r"C:\Users\Sunlight\Desktop\DOSSIE 09-06-2026 - Copie\3-AB-12-08-2024-TR"

files = list(Path(folder_path).rglob("*.xlsx"))

print("Total files in this folder:", len(files))
# ==========================================================
# 4) FINAL COLUMNS
# ==========================================================
COLUMNS = [
    "STE", "NAME", "TVA", "ADDRESS", "HOUSE_NUM", "CP", "CITY",
    "GSM", "TEL", "FORME_JURIDIQUE", "CODE_LINGUISTIQUE",
    "CIVILITE", "COMMENTS", "OTHER_INFO", "SOURCE_FILE"
]
# ==========================================================
# 5) RENAME MAP
# ==========================================================
rename_map = {
    "DENOMINATION SOCIALE": "STE",
    "DÉNOMINATION SOCIALE": "STE",
    "NAME_ORG1": "STE",
    "NAME_ORG2": "STE",
    "RS": "STE",
    "ENTITY NAME": "STE",
    "TITLE": "STE",

    "NOM": "NAME",
    "PRENOM": "NAME",
    "PRÉNOM": "NAME",
    "NAME": "NAME",

    "RUE": "ADDRESS",
    "ADRESSE": "ADDRESS",
    "STREET60": "ADDRESS",
    "ADRESSE COMPLET": "ADDRESS",
    "ADRES": "ADDRESS",
    "STRAAT": "ADDRESS",

    "NUMERO DE MAISON": "HOUSE_NUM",
    "NUMÉRO DE MAISON": "HOUSE_NUM",
    "AJOUT DE NUMERO DE MAISON": "HOUSE_NUM",
    "AJOUT DE NUMÉRO DE MAISON": "HOUSE_NUM",
    "NUMERO": "HOUSE_NUM",
    "NUMÉRO": "HOUSE_NUM",
    "HOUSE_NUM1": "HOUSE_NUM",
    "HOUSE_NUM2": "HOUSE_NUM",
    "NUM": "HOUSE_NUM",
    "N": "HOUSE_NUM",

    "CODE POSTAL": "CP",
    "POSTALCODE": "CP",
    "ZIPCODE": "CP",
    "PC": "CP",
    "POST CODE": "CP",

    "LIEU(COMPLET)": "CITY",
    "PROVINCE": "CITY",
    "CITY-1": "CITY",
    "CITY": "CITY",
    "VILLE": "CITY",

    "NUMERO D'ENTREPRISE": "TVA",
    "NUMÉRO D'ENTREPRISE": "TVA",
    "VAT_REG_NO": "TVA",
    "TVA": "TVA",
    "FINALVAT": "TVA",

    "GSM": "GSM",
    "PHONE_CLEAN": "GSM",
    "TEL2": "GSM",

    "NUMERO DE TELEPHONE": "TEL",
    "NUMÉRO DE TÉLÉPHONE": "TEL",
    "TELEPHONE": "TEL",
    "TÉLÉPHONE": "TEL",
    "PHONE": "TEL",
    "FIX": "TEL",
    "NO DE TELEPHONE": "TEL",
    "NO DE TÉLÉPHONE": "TEL",
    "FINAL PHONE": "TEL",

    "DESCRIPTION": "COMMENTS"
}
# ==========================================================
# 6) CLEANING FUNCTIONS
# ==========================================================
def clean_text(x):
    if pd.isna(x):
        return None
    x = str(x).strip()
    if x.lower() in ["nan", "none", ""]:
        return None
    return x


def clean_phone(x):
    if pd.isna(x):
        return None

    s = str(x).strip()

    if s.lower() in ["nan", "none", ""]:
        return None

    digits = re.sub(r"\D", "", s)

    if digits == "":
        return None

    # remove EAN/barcode/noise
    if len(digits) > 12:
        return None

    if digits.startswith("0032"):
        digits = "32" + digits[4:]

    elif digits.startswith("0"):
        digits = "32" + digits[1:]

    if digits.startswith("32") and len(digits) in [10, 11]:
        return digits

    if len(digits) in [8, 9]:
        return digits

    return None


def normalize_be_tva(x):
    """
    Keep TVA already like BE0462890433.
    """
    if pd.isna(x):
        return None

    s = str(x).upper().strip()
    s = re.sub(r"[^A-Z0-9]", "", s)

    if re.match(r"^BE[0-9]{9,10}$", s):
        digits = re.sub(r"\D", "", s)
        return "BE" + digits.zfill(10)

    return None


def normalize_numeric_tva_candidate(x):
    """
    Create TVA candidate from numeric values.
    We do NOT trust it yet. KBO must confirm it.
    """
    if pd.isna(x):
        return None

    s = str(x).strip()
    digits = re.sub(r"\D", "", s)

    if len(digits) == 10:
        return "BE" + digits

    if len(digits) == 9:
        return "BE" + digits.zfill(10)

    return None


def is_ean_noise(x):
    if pd.isna(x):
        return False
    digits = re.sub(r"\D", "", str(x))
    return len(digits) >= 13
# ==========================================================
# 7) STRICT KBO CHECK
# ==========================================================
session = requests.Session()

def format_kbo_number(number):
    """
    0462890433 -> 0462.890.433
    """
    number = str(number).zfill(10)
    return f"{number[:4]}.{number[4:7]}.{number[7:]}"


def check_tva_kbo_strict(tva):
    """
    Strict check:
    - Search KBO with enterprise number.
    - Accept only if returned page contains exact number in digit form
      OR formatted form like 0462.890.433.
    """
    if not tva:
        return False

    digits = re.sub(r"\D", "", str(tva).upper().replace("BE", ""))

    if len(digits) not in [9, 10]:
        return False

    digits10 = digits.zfill(10)
    formatted = format_kbo_number(digits10)

    url = "https://kbopub.economie.fgov.be/kbopub/zoeknummerform.html"

    try:
        response = session.get(
            url,
            params={"nummer": digits10},
            timeout=20,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Accept-Language": "fr-FR,fr;q=0.9,nl;q=0.8,en;q=0.7"
            }
        )

        html = response.text
        text = BeautifulSoup(html, "html.parser").get_text(" ", strip=True)

        text_digits = re.sub(r"\D", "", text)

        # Strict positive check
        if digits10 in text_digits:
            return True

        if formatted in text:
            return True

        # Negative words
        negative_words = [
            "aucun résultat",
            "geen resultaat",
            "no result",
            "niet gevonden",
            "introuvable"
        ]

        if any(word in text.lower() for word in negative_words):
            return False

        return False

    except Exception:
        return False
# ==========================================================
# 8) FILE CLASSIFICATION
# ==========================================================
def classify_file(file):
    try:
        df = pd.read_excel(file, nrows=5, engine="openpyxl")
        cols = [str(c).strip().upper() for c in df.columns]

        if any(col in rename_map for col in cols):
            return "TYPE_1_HEADER"

        return "TYPE_2_NO_HEADER"

    except:
        return "TYPE_2_NO_HEADER"
# ==========================================================
# 9) PROCESS TYPE 1: FILES WITH HEADERS
# ==========================================================
def process_type_1(file):
    df = pd.read_excel(file, engine="openpyxl")

    df.columns = [str(c).strip().upper() for c in df.columns]
    df.rename(columns=rename_map, inplace=True)

    df = df.groupby(level=0, axis=1).first()

    other_cols = [c for c in df.columns if c not in COLUMNS]

    if other_cols:
        df["OTHER_INFO"] = df[other_cols].apply(
            lambda row: " | ".join(
                f"{col}: {val}" for col, val in row.items()
                if pd.notna(val) and str(val).strip().lower() not in ["nan", ""]
            ),
            axis=1
        )
    else:
        df["OTHER_INFO"] = None

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None

    df["SOURCE_FILE"] = file.name

    return df[COLUMNS]
# ==========================================================
# 10) PROCESS TYPE 2: FILES WITHOUT HEADERS
# Format: STE | ADDRESS | CP | CITY | TEL | TVA
# ==========================================================
def process_type_2(file):
    df = pd.read_excel(file, header=None, engine="openpyxl")
    df = df.dropna(how="all")

    df = df.iloc[:, :6]

    while df.shape[1] < 6:
        df[df.shape[1]] = None

    df.columns = ["STE", "ADDRESS", "CP", "CITY", "TEL", "TVA"]

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None

    df["SOURCE_FILE"] = file.name

    return df[COLUMNS]
# ==========================================================
# 11) PROCESS ALL FILES IN ONE FOLDER
# ==========================================================
all_data = []
errors = []

for i, file in enumerate(files, start=1):
    try:
        file_type = classify_file(file)

        if file_type == "TYPE_1_HEADER":
            df = process_type_1(file)
        else:
            df = process_type_2(file)

        all_data.append(df)

        print(f"{i}/{len(files)} processed - {file_type} - {file.name}")

    except Exception as e:
        errors.append((file.name, str(e)))
        print("ERROR:", file.name, "->", e)

print("Processed files:", len(all_data))
print("Errors:", len(errors))
# ==========================================================
# 12) MERGE FOLDER DATA
# ==========================================================
final_df = pd.concat(all_data, ignore_index=True)

print("Rows before cleaning:", len(final_df))
print("Columns:", len(final_df.columns))
# ==========================================================
# 13) BASIC CLEANING
# ==========================================================
for col in ["STE", "NAME", "ADDRESS", "HOUSE_NUM", "CP", "CITY",
            "FORME_JURIDIQUE", "CODE_LINGUISTIQUE", "CIVILITE",
            "COMMENTS", "OTHER_INFO"]:
    final_df[col] = final_df[col].apply(clean_text)

print("Rows after text cleaning:", len(final_df))
# ==========================================================
# 14) REMOVE FAKE HEADER ROWS
# ==========================================================
bad_words = (
    "numero|numéro|denomination|dénomination|telephone|téléphone|"
    "adresse|postal|vat|tva|société|societe"
)

final_df = final_df[
    ~final_df["STE"].astype(str).str.lower().str.contains(bad_words, na=False)
]

print("Rows after removing fake headers:", len(final_df))
# ==========================================================
# 15) KEEP TVA THAT ALREADY STARTS WITH BE
# ==========================================================
final_df["TVA_BE_CLEAN"] = final_df["TVA"].apply(normalize_be_tva)

print("TVA already BE:", final_df["TVA_BE_CLEAN"].notna().sum())
print("TVA missing/suspicious:", final_df["TVA_BE_CLEAN"].isna().sum())
# ==========================================================
# 16) CREATE STRICT CANDIDATES ONLY FOR SUSPICIOUS ROWS
# Search only TVA, GSM, TEL, OTHER_INFO.
# ==========================================================
def find_candidates_for_row(row):
    candidates = []

    for col in ["TVA", "GSM", "TEL", "OTHER_INFO"]:
        candidate = normalize_numeric_tva_candidate(row.get(col))

        if candidate:
            candidates.append(candidate)

    return list(dict.fromkeys(candidates))


suspicious_mask = final_df["TVA_BE_CLEAN"].isna()

final_df.loc[suspicious_mask, "TVA_CANDIDATES"] = final_df.loc[suspicious_mask].apply(
    lambda row: ", ".join(find_candidates_for_row(row)),
    axis=1
)

final_df.loc[~suspicious_mask, "TVA_CANDIDATES"] = ""

print("Suspicious rows:", suspicious_mask.sum())
print("Rows with candidates:", (final_df["TVA_CANDIDATES"] != "").sum())
# ==========================================================
# 17) UNIQUE CANDIDATES ONLY
# ==========================================================
unique_candidates = (
    final_df.loc[suspicious_mask, "TVA_CANDIDATES"]
    .dropna()
    .astype(str)
    .str.split(", ")
    .explode()
    .str.strip()
    .drop_duplicates()
)

unique_candidates = unique_candidates[unique_candidates != ""]

print("Unique candidates to check:", len(unique_candidates))
print(unique_candidates.head(20))
# ==========================================================
# 18) TEST STRICT KBO ON FIRST 20
# Run this before checking all.
# ==========================================================
test_results = []

for i, tva in enumerate(unique_candidates.head(20), start=1):
    valid = check_tva_kbo_strict(tva)
    test_results.append((tva, valid))
    print(i, tva, valid)
    time.sleep(0.7)

test_df = pd.DataFrame(test_results, columns=["TVA_CANDIDATE", "IS_VALID_KBO"])
test_df
# ==========================================================
# 19) CHECK ALL UNIQUE CANDIDATES
# ONLY run if Step 18 works correctly.
# ==========================================================
checked_results = []

for i, tva in enumerate(unique_candidates, start=1):
    valid = check_tva_kbo_strict(tva)
    checked_results.append((tva, valid))

    if i % 100 == 0:
        print(i, "/", len(unique_candidates))

    time.sleep(0.7)

checked_df = pd.DataFrame(
    checked_results,
    columns=["TVA_CANDIDATE", "IS_VALID_KBO"]
)

checked_df.to_excel("kbo_checked_candidates_one_folder.xlsx", index=False)

print("KBO checking done")
# ==========================================================
# 20) APPLY KBO RESULTS
# ==========================================================
valid_tvas = set(
    checked_df.loc[checked_df["IS_VALID_KBO"], "TVA_CANDIDATE"]
)

def choose_final_tva(row):
    # keep BE if already clean
    if pd.notna(row["TVA_BE_CLEAN"]):
        return row["TVA_BE_CLEAN"]

    candidates = str(row.get("TVA_CANDIDATES", "")).split(", ")

    for candidate in candidates:
        if candidate in valid_tvas:
            return candidate

    return None


final_df["TVA"] = final_df.apply(choose_final_tva, axis=1)

print("Final valid TVA:", final_df["TVA"].notna().sum())
# ==========================================================
# 21) CLEAN GSM/TEL AFTER TVA DECISION
# ==========================================================
final_df["GSM"] = final_df["GSM"].apply(clean_phone)
final_df["TEL"] = final_df["TEL"].apply(clean_phone)
# ==========================================================
# 22) REMOVE EAN NOISE
# ==========================================================
final_df["OTHER_INFO"] = final_df["OTHER_INFO"].apply(
    lambda x: None if is_ean_noise(x) else x
)

final_df = final_df[
    final_df["STE"].notna() |
    final_df["TVA"].notna() |
    final_df["GSM"].notna() |
    final_df["TEL"].notna()
]

print("Rows after EAN/noise cleaning:", len(final_df))
# ==========================================================
# 23) DEDUPLICATE
# ==========================================================
with_tva = final_df[final_df["TVA"].notna()].copy()
without_tva = final_df[final_df["TVA"].isna()].copy()

with_tva = with_tva.drop_duplicates(subset=["TVA"], keep="first")

without_tva = without_tva.drop_duplicates(
    subset=["STE", "CP", "CITY"],
    keep="first"
)

final_df = pd.concat([with_tva, without_tva], ignore_index=True)

print("Rows after deduplication:", len(final_df))
# ==========================================================
# 24) DROP TEMP COLUMNS
# ==========================================================
final_df = final_df.drop(
    columns=["TVA_BE_CLEAN", "TVA_CANDIDATES"],
    errors="ignore"
)
# ==========================================================
# 25) EXPORT ONE FOLDER RESULT
# ==========================================================
output_file = "master_one_folder_clean_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"

final_df.to_excel(output_file, index=False, engine="openpyxl")

print("DONE:", output_file)
# ==========================================================
# 26) CHECK RESULT
# ==========================================================
final_df[["STE", "TVA", "GSM", "TEL", "SOURCE_FILE"]].head(30)


