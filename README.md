def classify_file(file):
    try:
        df = pd.read_excel(file, nrows=5, engine="openpyxl")
        cols = [str(c).strip().upper() for c in df.columns]

        if any(col in rename_map for col in cols):
            return "TYPE_1_HEADER"

        return "TYPE_2_NO_HEADER"

    except:
        return "TYPE_2_NO_HEADER"
