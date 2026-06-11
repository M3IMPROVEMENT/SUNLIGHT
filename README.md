unique_candidates = (
    final_df["TVA_CANDIDATES"]
    .dropna()
    .astype(str)
    .str.split(", ")
    .explode()
    .str.strip()
    .drop_duplicates()
)

unique_candidates = unique_candidates[unique_candidates != ""]

print("Unique candidates:", len(unique_candidates))

print(unique_candidates.head(30))
