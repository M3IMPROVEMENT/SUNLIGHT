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
