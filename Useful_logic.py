# Derived-column logic (Single derived column vs Multiple derived columns
"""
1. Read one or more existing columns (Which original columns does this depend on?)
2. Compute a new value (Read all needed values BEFORE inserting columns and store them in memory (dict by row))
3. Insert new column(s) (Insert columns only after caching and write headers)
4. Fill values row-by-row (Loop rows → compute → write values)
5. Avoid breaking column indices

Cached version:
-Use for single derived columns 
-Lower memory usage 
Disadvantages:
-Reads from excel every time
-Much slower on larger files 
-Struggles to handle multiple columns
-Sensitive to column shifts (order changes)
-Struggles also when logic grows
vs 
Non-cached version 
-Works faster on larger files
-Reads from excel once
-Not sensitive to column shifts
-Able to handle complex transformations if logic scales
Disadvantages:
-Higher memory usage 
-Conceptually more complex

"""


#Single derived column (Uses non-cached approach) 
def add_ft_pt_column(ws, headers):
    """
    Adds a column 'FT/PT' next to 'Usual hours of work'
    FT if hours >= 35
    PT if hours < 35
    """

    usual_hours_header = normalise_header("Usual hours of work")

    if usual_hours_header not in headers:
        return  # header not found → do nothing

 # IMPORTANT:
    # Iterate in REVERSE order so column insertion doesn't shift later indices
    for hours_col in sorted(headers[usual_hours_header], reverse=True):

        # Step 1: check if there is at least ONE value below this column
        has_any_value = False
        for row in range(7, ws.max_row + 1):
            if ws.cell(row=row, column=hours_col).value not in (None, ""):
                has_any_value = True
                break

        if not has_any_value:
            continue  # skip completely empty hours column

        # Step 2: insert FT/PT column beside it
        ft_pt_col = hours_col + 1
        ws.insert_cols(ft_pt_col)


        # Step 3: Write header
        ws.cell(row=6, column=ft_pt_col).value = "FT/PT"

        # Fill values
        for row in range(7, ws.max_row + 1):
            value = ws.cell(row=row, column=hours_col).value

            if value in (None, ""):
                continue

            try:
                hours = float(value)
            except ValueError:
                continue

            if hours >= 35:
                ws.cell(row=row, column=ft_pt_col).value = "FT"
            else:
                ws.cell(row=row, column=ft_pt_col).value = "PT"

#Multiple derived column (Uses cached approach)
def add_professional_certification_columns(ws, headers):

    base_question = normalise_header(
        "Have you ever obtained any Vocational or Skills certificates/qualifications, "
        "(e.g. (WSQ) and (ESS) certificates, or formal certifications that validate "
        "knowledge and skills in a particular field)?"
    )

    certification_header = normalise_header("Professional Certification")

    certification_options = [
        "Care Economy",
        "Artificial Intelligence",
        "Digital Skills",
        "Green Economy",
        "Industry 4.0"
    ]

    if base_question not in headers or certification_header not in headers:
        return

    base_col = headers[base_question][0]
    cert_col = headers[certification_header][0]

    # 🔹 Cache certification values BEFORE column insertion
    cert_values_by_row = {
        row: ws.cell(row=row, column=cert_col).value
        for row in range(7, ws.max_row + 1)
    }

    insert_start_col = base_col + 1
    ws.insert_cols(insert_start_col, amount=len(certification_options))

    # Write headers
    for i, option in enumerate(certification_options):
        ws.cell(row=6, column=insert_start_col + i).value = option

    # Populate values
    for row in range(7, ws.max_row + 1):

        base_value = ws.cell(row=row, column=base_col).value
        cert_value = cert_values_by_row.get(row)

        if base_value is None or str(base_value).strip() == "":
            continue

        base_norm = normalise(base_value)

        if base_norm not in ("yes", "no"):
            continue

        selected_options = set()
        if cert_value:
            parts = str(cert_value).replace(";", ",").split(",")
            selected_options = {normalise(p) for p in parts if p.strip()}

        for i, option in enumerate(certification_options):
            col = insert_start_col + i

            if base_norm == "no":
                ws.cell(row=row, column=col).value = "No"
            else:
                ws.cell(row=row, column=col).value = (
                    "Yes" if normalise(option) in selected_options else "No"
                )

