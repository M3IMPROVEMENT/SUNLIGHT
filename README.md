final_df = pd.concat(all_data, ignore_index=True)

print("Rows before cleaning:", len(final_df))


for col in ["STE", "NAME", "ADDRESS", "HOUSE_NUM", "CP", "CITY",
            "FORME_JURIDIQUE", "CODE_LINGUISTIQUE", "CIVILITE",
            "COMMENTS", "OTHER_INFO"]:
    final_df[col] = final_df[col].apply(clean_text)
