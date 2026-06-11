files = (
    list(Path(r"C:\Users\Sunlight\Desktop\DOSSIE 09-06-2026 - Copie\3-AB-12-08-2024-TR").rglob("*.xlsx"))
    +
    list(Path(r"C:\Users\Sunlight\Desktop\DOSSIE 09-06-2026 - Copie\Chaima 10-06-2024--TR").rglob("*.xlsx"))
    +
    list(Path(r"C:\Users\Sunlight\Desktop\DOSSIE 09-06-2026 - Copie\Imane 10-06-2024--TR").rglob("*.xlsx"))
)

print("Total files:", len(files))



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
