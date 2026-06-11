good_tva = final_df[
    final_df["TVA"].astype(str).str.match(r"^[A-Z]{2}", na=False)
]

bad_tva = final_df[
    ~final_df["TVA"].astype(str).str.match(r"^[A-Z]{2}", na=False)
]

print("Good TVA:", len(good_tva))
print("Bad TVA:", len(bad_tva))
