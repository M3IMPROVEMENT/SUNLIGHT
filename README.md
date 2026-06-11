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
