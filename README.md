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
