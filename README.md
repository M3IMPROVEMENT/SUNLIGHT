unique_candidates = (
    final_df["TVA_CANDIDATES"]
    .dropna()
    .astype(str)
    .str.split(", ")
    .explode()
    .dropna()
    .drop_duplicates()
)

unique_candidates = unique_candidates[unique_candidates != ""]

print("Unique TVA candidates:", len(unique_candidates))
