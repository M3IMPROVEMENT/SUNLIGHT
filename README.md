def check_tva_kbo(tva):
    """
    Checks a Belgian company number on KBO Public Search.
    Returns True if the number appears valid/found, False otherwise.

    Important: use this only for suspicious candidates, not millions of rows.
    """
    if not tva:
        return False

    number = str(tva).upper().replace("BE", "")
    number = re.sub(r"\D", "", number)

    if len(number) not in [9, 10]:
        return False

    url = "https://kbopub.economie.fgov.be/kbopub/zoeknummerform.html"

    try:
        response = requests.get(
            url,
            params={"nummer": number},
            timeout=15,
            headers={
                "User-Agent": "Mozilla/5.0"
            }
        )

        text = response.text.lower()

        # If the number appears in returned page, likely found
        if number in re.sub(r"\D", "", text):
            return True

        invalid_words = [
            "geen resultaat",
            "aucun résultat",
            "no result",
            "niet gevonden",
            "introuvable"
        ]

        if any(word in text for word in invalid_words):
            return False

        return False

    except Exception:
        return False
