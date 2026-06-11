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
