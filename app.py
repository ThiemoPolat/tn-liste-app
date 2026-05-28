import re
from io import BytesIO
from pathlib import Path

import streamlit as st
from docx import Document


st.set_page_config(
    page_title="TN-Liste in Word-Vorlage einfügen",
    page_icon="📄",
    layout="centered",
)

st.title("TN-Liste in Word-Vorlage einfügen")
st.write(
    "Word-Vorlage hochladen, TN-Liste einfügen und die vorhandene Tabelle automatisch befüllen lassen."
)


def parse_tn_list(raw_text: str):
    """
    Erwartetes Format pro Teilnehmer:
    Kürzel
    Vollständiger Name
    Dokumentnummer
    Status

    Beispiel:
    GK
    Ginette Kouowou
    D384291
    Initial
    """
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    participants = []

    i = 0
    while i < len(lines):
        # Suche nach Kürzel + Name + Dokumentnummer
        if (
            i + 2 < len(lines)
            and re.fullmatch(r"[A-Za-zÄÖÜäöüß]{1,5}", lines[i])
            and re.fullmatch(r"D\d+", lines[i + 2])
        ):
            name = lines[i + 1]
            participants.append(name)
            # meistens folgt danach Status wie Initial / In Lobby
            i += 4
        else:
            i += 1

    return participants


def make_pairs(participants):
    pairs = []
    for i in range(0, len(participants), 2):
        tn1 = participants[i]
        tn2 = participants[i + 1] if i + 1 < len(participants) else ""
        if tn2:
            pairs.append(f"{tn1} (1)\n{tn2} (2)")
        else:
            pairs.append(f"{tn1} (1)")
    return pairs


def clear_cell_keep_first_paragraph(cell):
    """
    Entfernt Zellinhalt möglichst schonend.
    Die Tabellenstruktur bleibt erhalten.
    """
    for p in cell.paragraphs:
        p.clear()
    if not cell.paragraphs:
        cell.add_paragraph()
    return cell.paragraphs[0]


def fill_cell_with_lines(cell, text):
    """
    Füllt eine vorhandene Tabellenzelle, ohne die Tabelle neu zu erzeugen.
    """
    paragraph = clear_cell_keep_first_paragraph(cell)
    lines = text.split("\n")

    for idx, line in enumerate(lines):
        run = paragraph.add_run(line)
        if idx < len(lines) - 1:
            run.add_break()


def fill_template(docx_file, pairs):
    doc = Document(docx_file)

    if not doc.tables:
        raise ValueError("In der Vorlage wurde keine Tabelle gefunden.")

    table = doc.tables[0]

    # Zeilen mit Pause erkennen und Teilnehmer in Spalte 2 einfügen.
    pair_index = 0
    for row in table.rows[1:]:
        if len(row.cells) < 2:
            continue

        row_text = " ".join(cell.text.strip() for cell in row.cells)

        if "PAUSE" in row_text.upper():
            fill_cell_with_lines(row.cells[1], "PAUSE")
            continue

        if pair_index < len(pairs):
            fill_cell_with_lines(row.cells[1], pairs[pair_index])
            pair_index += 1

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output, pair_index


uploaded_template = st.file_uploader("Word-Vorlage hochladen (.docx)", type=["docx"])

tn_text = st.text_area(
    "TN-Liste hier einfügen",
    height=360,
    placeholder="GK\nGinette Kouowou\nD384291\nInitial\nPM\nPamela Marthe Huguette MBACK\nD384295\nInitial",
)

output_name = st.text_input(
    "Dateiname für die fertige Datei",
    value="TN_Liste_befuellt.docx",
)

if st.button("Vorlage befüllen", type="primary"):
    if uploaded_template is None:
        st.error("Bitte zuerst eine Word-Vorlage hochladen.")
    elif not tn_text.strip():
        st.error("Bitte die TN-Liste einfügen.")
    else:
        try:
            participants = parse_tn_list(tn_text)
            pairs = make_pairs(participants)

            if not participants:
                st.error("Es konnten keine Teilnehmer erkannt werden. Bitte Format prüfen.")
            else:
                result, inserted_pairs = fill_template(uploaded_template, pairs)

                st.success(
                    f"Fertig. Erkannt: {len(participants)} Teilnehmer, eingefügt: {inserted_pairs} Paar(e)."
                )

                if len(participants) % 2 == 1:
                    st.warning("Ungerade Teilnehmerzahl: Der letzte Teilnehmer wurde allein als (1) eingefügt.")

                st.download_button(
                    label="Fertiges Word-Dokument herunterladen",
                    data=result,
                    file_name=output_name if output_name.endswith(".docx") else output_name + ".docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

        except Exception as e:
            st.error(f"Fehler beim Befüllen der Vorlage: {e}")