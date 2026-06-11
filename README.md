def validate_tva_with_kbo_or_row(row):
    # 1) Try TVA column first
    main_tva = normalize_tva_candidate(row.get("TVA"))

    if main_tva and check_tva_kbo(main_tva):
        return main_tva

    # 2) If TVA column failed, search all row values
    candidates = find_tva_candidates_in_row(row)

    for candidate in candidates:
        if check_tva_kbo(candidate):
            return candidate

        time.sleep(0.4)

    # 3) Nothing valid found
    return None
