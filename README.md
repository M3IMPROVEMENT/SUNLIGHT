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

print("Total unique candidates:", len(unique_candidates))

numeric_candidates = unique_candidates[
    unique_candidates.str.match(r"^BE[0-9]{10}$", na=False)
]

print("BE-format candidates:", len(numeric_candidates))
