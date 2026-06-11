valid_format = unique_tvas[
    unique_tvas.str.match(r"^[A-Z]{2}[0-9]+$", na=False)
]

numeric_only = unique_tvas[
    unique_tvas.str.match(r"^[0-9]{9,10}$", na=False)
]

other = unique_tvas[
    ~unique_tvas.isin(valid_format) &
    ~unique_tvas.isin(numeric_only)
]

print("Already valid format:", len(valid_format))
print("Numeric only:", len(numeric_only))
print("Other / suspicious:", len(other))
