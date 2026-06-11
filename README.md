def find_tva_candidates_in_row(row):
    candidates = []

    for value in row.values:
        candidate = normalize_tva_candidate(value)
        if candidate:
            candidates.append(candidate)

    return list(dict.fromkeys(candidates))


def choose_best_tva(row):
    """
    Priority:
    1. TVA column if valid and KBO confirms
    2. Any candidate in the row if KBO confirms
    3. Return None if no confirmed TVA
    """

    candidates = []

    # Priority: current TVA column first
    if "TVA" in row.index:
        first = normalize_tva_candidate(row["TVA"])
        if first:
            candidates.append(first)

    # Then scan whole row
    for candidate in find_tva_candidates_in_row(row):
        if candidate not in candidates:
            candidates.append(candidate)

    for candidate in candidates:
        if check_tva_kbo(candidate):
            time.sleep(0.4)
            return candidate

        time.sleep(0.4)

    return None
