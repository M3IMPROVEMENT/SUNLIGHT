be_candidates = unique_candidates[
    unique_candidates.str.match(r"^BE[0-9]{10}$", na=False)
]

print("BE candidates:", len(be_candidates))
