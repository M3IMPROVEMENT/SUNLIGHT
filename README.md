final_df = pd.concat(all_data, ignore_index=True)

print("Rows before cleaning:", len(final_df))


for col in ["STE", "NAME", "ADDRESS", "HOUSE_NUM", "CP", "CITY",
            "FORME_JURIDIQUE", "CODE_LINGUISTIQUE", "CIVILITE",
            "COMMENTS", "OTHER_INFO"]:
    final_df[col] = final_df[col].apply(clean_text)



bad_words = (
    "numero|numéro|denomination|dénomination|telephone|téléphone|"
    "adresse|postal|vat|tva|société|societe"
)

final_df = final_df[
    ~final_df["STE"].astype(str).str.lower().str.contains(bad_words, na=False)
]

print("Rows after removing fake headers:", len(final_df))



final_df["GSM"] = final_df["GSM"].apply(clean_phone)
final_df["TEL"] = final_df["TEL"].apply(clean_phone)




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




