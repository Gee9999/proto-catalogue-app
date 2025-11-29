def match_photos_to_prices(photo_files, price_df):

    # Normalize CODE column
    price_df["CODE_STR"] = price_df["CODE"].astype(str)

    # FIX: remove duplicates to avoid "index must be unique"
    price_df = price_df.drop_duplicates(subset="CODE_STR", keep="first")

    # Build lookup
    price_dict = price_df.set_index("CODE_STR")[["DESCRIPTION", "PRICE_A_INCL"]].to_dict("index")

    rows = []
    for photo in photo_files:
        fname = photo.name.upper()
        numeric = "".join(c for c in fname if c.isdigit())
        code_str = numeric[:10]

        if code_str in price_dict:
            info = price_dict[code_str]
            rows.append({
                "PHOTO_FILE": photo,
                "CODE": code_str,
                "DESCRIPTION": info["DESCRIPTION"],
                "PRICE_A_INCL": info["PRICE_A_INCL"]
            })
        else:
            rows.append({
                "PHOTO_FILE": photo,
                "CODE": code_str,
                "DESCRIPTION": "",
                "PRICE_A_INCL": ""
            })

    return pd.DataFrame(rows)
