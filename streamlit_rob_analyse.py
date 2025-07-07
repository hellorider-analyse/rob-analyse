# -*- coding: utf-8 -*-
"""
Created on Mon Jul  7 15:42:35 2025

@author: MijkePietersma
"""

import re
import streamlit as st
import pandas as pd


st.title("ROB‑analyse")

# 1. uploads …
admin1_0700_file = st.file_uploader("0700 – Admin 1", type="xlsx", key="a1_0700")
admin2_0700_file = st.file_uploader("0700 – Admin 2", type="xlsx", key="a2_0700")
admin1_5000_file = st.file_uploader("5000 – Admin 1", type="xlsx", key="a1_5000")
metabase_file    = st.file_uploader("Metabase export", type="xlsx", key="metabase")
alle_projecten_file = st.file_uploader("Alle projecten HR", type="xlsx", key="projects")

uploads_ready = all(f is not None for f in (
    admin1_0700_file, admin2_0700_file, admin1_5000_file, metabase_file, alle_projecten_file
))

if uploads_ready:
    run = st.button("Run analyse")
else:
    st.info("⬆️  Upload alle vijf de Excel‑bestanden.")
    run = False

# ------------------------------------------------------------ #
# 2.  Analyse & export
# ------------------------------------------------------------ #
if run:
    with st.spinner("Bestanden lezen…"):
        admin1_0700 = pd.read_excel(admin1_0700_file)
        admin2_0700 = pd.read_excel(admin2_0700_file)
        admin1_5000 = pd.read_excel(admin1_5000_file)

        SHEET = "Query result"
        header_df = pd.read_excel(metabase_file, sheet_name=SHEET, nrows=0)
        proj_col = header_df.columns[2]

        metabase_file.seek(0)
        metabase = pd.read_excel(metabase_file, sheet_name=SHEET,
                                 dtype={proj_col: str})

        alle_projecten = pd.read_excel(alle_projecten_file)

    # --------------------------------------------------------------------------- #
    # 4.  Kolommen ‘Admin’ en ‘Rek’ toevoegen
    # --------------------------------------------------------------------------- #
    for df_, admin, rek in [
        (admin1_0700, 1, "0700"),
        (admin2_0700, 2, "0700"),
        (admin1_5000, 1, "5000"),
    ]:
        df_.insert(0, "Admin", admin)
        df_.insert(1, "Rek", rek)

    # --------------------------------------------------------------------------- #
    # 5.  Drie tabellen samenvoegen
    # --------------------------------------------------------------------------- #
    df = pd.concat([admin1_0700, admin2_0700, admin1_5000], ignore_index=True)

    # --------------------------------------------------------------------------- #
    # 6.  Kolom ‘Verkooprelatie’
    # --------------------------------------------------------------------------- #
    project_map = dict(
        zip(alle_projecten["Projectnummer"], alle_projecten["Verkooprelatie"])
    )
    df["Verkooprelatie"] = df["Code verbijzonderingsas 1 Verb. 1"].map(project_map)

    na_mask = df["Verkooprelatie"].isna()
    df.loc[na_mask & (df["Code verbijzonderingsas 1 Verb. 1"] == "*****"),
           "Verkooprelatie"] = "Verkooprelatie"
    df.loc[na_mask & (df["Code verbijzonderingsas 1 Verb. 1"] == "CORR"),
           "Verkooprelatie"] = "Arval Correctie"

    # …op plek 4 zetten
    cols = df.columns.tolist()
    cols.insert(3, cols.pop(cols.index("Verkooprelatie")))
    df = df[cols]

    # --------------------------------------------------------------------------- #
    # 7.  Kolom ‘Incl/Excl’
    # --------------------------------------------------------------------------- #
    lease_partners = [
        "BVD Lease B.V", "Mobilease E-bikeleasing B.V", "Wittebrug Lease B.V",
        "Fiets Lease Partner B", "Dynamo Lease B.V", "Arval Correctie",
    ]
    partner_pattern = "|".join(map(re.escape, lease_partners))

    df["Incl/Excl"] = df["Verkooprelatie"].str.contains(
        partner_pattern, case=False, na=False
    ).map({True: "Excl", False: "Incl"})

    cols = df.columns.tolist()
    cols.insert(4, cols.pop(cols.index("Incl/Excl")))
    df = df[cols]

    # --------------------------------------------------------------------------- #
    # 8.  Kolom ‘Incl/Excl2’
    # --------------------------------------------------------------------------- #
    boeking_patterns = [
        r"^\[(?:vrk|bdf|afs)\]",                       r"^Afboek",
        r"^Corr,?\s*(?:diefstal|total loss|verkoop|Urban Arrow|onterecht)",
        r"^Corr\.?\s*(?:CBRE|kosten reparatie|leasefiets|Stromer)",
        r"^Omzet ROB",                                r"corr arval rob",
        r"^Gesloten project",                         r"Voorziening|Vrijval|Leegboeken",
        r"^ROB arval",                                r"afboeken crediteur",
        r"totaal voorzien",
    ]
    boeking_re = re.compile("|".join(boeking_patterns), flags=re.IGNORECASE)

    def incl_excl2(boeking: str | float | None) -> str:
        if pd.isna(boeking):
            return "Excl"
        return "Excl" if boeking_re.search(str(boeking)) else "Incl"

    df["Incl/Excl2"] = df["Boeking"].apply(incl_excl2)

    cols = df.columns.tolist()
    cols.insert(13, cols.pop(cols.index("Incl/Excl2")))
    df = df[cols]

    # --------------------------------------------------------------------------- #
    # 9.  Alleen Incl‑Incl‑regels selecteren
    # --------------------------------------------------------------------------- #
    incl_df = df[(df["Incl/Excl"] == "Incl") & (df["Incl/Excl2"] == "Incl")]

    # --------------------------------------------------------------------------- #
    # 10.  Pivottabellen
    # --------------------------------------------------------------------------- #
    project_id_col = df.columns[2]          # derde kolom = Projectnummer
    project_sums = (
        incl_df.groupby(project_id_col, dropna=False)["Saldo"]
               .sum().reset_index()
               .rename(columns={project_id_col: "ProjectID", "Saldo": "AFAS_Saldo"})
    )

    metabase_amount_col  = metabase.columns[1]
    metabase_project_col = metabase.columns[2]
    metabase_sums = (
        metabase.groupby(metabase_project_col, dropna=False)[metabase_amount_col]
                .sum().reset_index()
                .rename(columns={metabase_project_col: "ProjectID",
                                 metabase_amount_col: "Metabase_Saldo"})
    )

    # --------------------------------------------------------------------------- #
    # 11.  Samenvoegen + verschilkolom
    # --------------------------------------------------------------------------- #
    combined = project_sums.merge(metabase_sums, on="ProjectID", how="left")
    combined["Metabase_Saldo"] = combined["Metabase_Saldo"].fillna(0)
    combined["Difference"]     = (combined["AFAS_Saldo"]
                                  - combined["Metabase_Saldo"]).round(2)

    only_AFAS     = combined[(combined["Metabase_Saldo"] == 0)
                             & (combined["Difference"] != 0)]
    only_Metabase = combined[(combined["AFAS_Saldo"] == 0)
                             & (combined["Difference"] != 0)]
    AFAS_Metabase = combined[(combined["Difference"] != 0)
                             & (combined["AFAS_Saldo"]     != 0)
                             & (combined["Metabase_Saldo"] != 0)]

    # --------------------------------------------------------------------------- #
    # 12.  Excel‑export
    # --------------------------------------------------------------------------- #
        
    from io import BytesIO
    
    # … je analyse-code hierboven …
    
    output = BytesIO()
    
    # 1) Schrijf met xlsxwriter + constant_memory
    with st.spinner("Output wordt gegenereerd..."):
        with pd.ExcelWriter(
            output,
            engine="xlsxwriter"
        ) as xl:
            df.to_excel(xl,              sheet_name="AFAS data",          index=False)
            metabase.to_excel(xl,        sheet_name="Metabase data",      index=False)
            combined.to_excel(xl,        sheet_name="Alle AFAS‑Metabase", index=False)
            only_AFAS.to_excel(xl,       sheet_name="Wel AFAS",           index=False)
            only_Metabase.to_excel(xl,   sheet_name="Wel Metabase",       index=False)
            AFAS_Metabase.to_excel(xl,   sheet_name="Verschillen",        index=False)
    
    # 2) Rewind en haal ruwe bytes op
    output.seek(0)
    bytes_data = output.getvalue()
    
    # 3) Download‑knop
    st.download_button(
        label="Download resultaat‑Excel",
        data=bytes_data,   # <-- niet het BytesIO‑object zelf
        file_name="ROB_analyse.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    
    st.success("✅ Analyse voltooid – bestand klaar voor download.")








