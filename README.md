test_df = final_df.head(10).copy()

test_df["TVA_CHECKED"] = test_df.apply(choose_best_tva, axis=1)

test_df[["STE", "TVA", "GSM", "TEL", "TVA_CANDIDATES", "TVA_CHECKED"]]




final_df["TVA_CHECKED"] = final_df.apply(choose_best_tva, axis=1)

final_df["TVA"] = final_df["TVA_CHECKED"]

final_df.drop(columns=["TVA_CHECKED"], inplace=True)

print("Valid checked TVA:", final_df["TVA"].notna().sum())


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



with_tva = final_df[final_df["TVA"].notna()].copy()
without_tva = final_df[final_df["TVA"].isna()].copy()

with_tva = with_tva.drop_duplicates(subset=["TVA"], keep="first")

without_tva = without_tva.drop_duplicates(
    subset=["STE", "CP", "CITY"],
    keep="first"
)

final_df = pd.concat([with_tva, without_tva], ignore_index=True)

print("Rows after deduplication:", len(final_df))



output_file = "master_companies_checked_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"

final_df.to_excel(
    output_file,
    index=False,
    engine="openpyxl"
)

print("DONE:", output_file)


