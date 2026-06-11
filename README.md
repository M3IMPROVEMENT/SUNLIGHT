# Split all TVA candidates from all rows
all_candidates = (
    final_df["TVA_CANDIDATES"]
    .dropna()
    .astype(str)
    .str.split(", ")
    .explode()
    .str.strip()
)

all_candidates = all_candidates[all_candidates != ""]

print("Total candidates:", len(all_candidates))

# Count candidates that start with BE
starts_be = all_candidates.str.startswith("BE")

print("Start with BE:", starts_be.sum())
print("Do NOT start with BE:", (~starts_be).sum())
