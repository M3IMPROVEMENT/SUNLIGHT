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
