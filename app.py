import streamlit as st
import calendar
import datetime
import io
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import pandas as pd

# ─── Jours fériés France ──────────────────────────────────────────────────────
def jours_feries(annee):
    from workalendar.europe import France
    cal = France()
    return {d for d, _ in cal.holidays(annee)}

JOURS_FR = ['Lun.', 'Mar.', 'Mer.', 'Jeu.', 'Ven.', 'Sam.', 'Dim.']
MOIS_FR  = ['', 'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
            'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']

def get_jours_mois(annee, mois):
    feries = jours_feries(annee)
    nb_jours = calendar.monthrange(annee, mois)[1]
    jours = []
    for j in range(1, nb_jours + 1):
        date = datetime.date(annee, mois, j)
        dow  = date.weekday()
        est_ferie    = date in feries
        est_samedi   = (dow == 5)
        est_dimanche = (dow == 6)
        label = f"{JOURS_FR[dow]} {j:02d}/{mois:02d}/{annee}"
        if est_ferie:
            label = f"Férié {j:02d}/{mois:02d}/{annee}"
        jours.append({
            'date': date, 'label': label,
            'samedi': est_samedi, 'dimanche': est_dimanche,
            'ferie': est_ferie,
            'weekend': est_samedi or est_dimanche,
        })
    return jours

# ─── Constantes calquées sur le modèle Node.js original ──────────────────────
DATE_W_DXA  = 2600
HEURE_W_DXA = 1400
FONT        = 'Arial'

COLOR_HEADER = '1F1F1F'
COLOR_DIM    = 'AAAAAA'
COLOR_FERIE  = 'DDDDDD'
COLOR_WHITE  = 'FFFFFF'

# ─── Helpers XML ──────────────────────────────────────────────────────────────
def _tcPr(cell):
    return cell._tc.get_or_add_tcPr()

def set_shading(cell, fill):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  fill)
    _tcPr(cell).append(shd)

def set_border(cell):
    tcBorders = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'),   'single')
        b.set(qn('w:sz'),    '4')
        b.set(qn('w:color'), '000000')
        tcBorders.append(b)
    _tcPr(cell).append(tcBorders)

def set_no_border(cell):
    tcBorders = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'),   'none')
        b.set(qn('w:sz'),    '0')
        b.set(qn('w:color'), 'FFFFFF')
        tcBorders.append(b)
    _tcPr(cell).append(tcBorders)

def set_cell_width(cell, dxa):
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'),    str(dxa))
    tcW.set(qn('w:type'), 'dxa')
    _tcPr(cell).insert(0, tcW)

def set_cell_margins(cell, top=60, bottom=60, left=120, right=120):
    tcMar = OxmlElement('w:tcMar')
    for side, val in (('top', top), ('bottom', bottom),
                      ('left', left), ('right', right)):
        m = OxmlElement(f'w:{side}')
        m.set(qn('w:w'),    str(val))
        m.set(qn('w:type'), 'dxa')
        tcMar.append(m)
    _tcPr(cell).append(tcMar)

def set_row_height(row, val_dxa, rule='exact'):
    trPr = row._tr.get_or_add_trPr()
    trH  = OxmlElement('w:trHeight')
    trH.set(qn('w:val'),   str(val_dxa))
    trH.set(qn('w:hRule'), rule)
    trPr.append(trH)

def fix_table_widths(table):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'),    str(DATE_W_DXA + HEURE_W_DXA))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)

    old = tbl.find(qn('w:tblGrid'))
    if old is not None:
        tbl.remove(old)
    tblGrid = OxmlElement('w:tblGrid')
    for w in (DATE_W_DXA, HEURE_W_DXA):
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(w))
        tblGrid.append(gc)
    tbl.insert(1, tblGrid)

    for row in table.rows:
        widths = [DATE_W_DXA, HEURE_W_DXA]
        for i, cell in enumerate(row.cells):
            set_cell_width(cell, widths[i])

# ─── Construction d'une fiche ─────────────────────────────────────────────────
def generer_fiche(doc, societe, nom, prenom, annee, mois,
                  date_entree=None, premier=True):
    if not premier:
        doc.add_page_break()

    def para_centre(text, bold=False, size=11, after=80):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        pPr = p._p.get_or_add_pPr()
        sp  = OxmlElement('w:spacing')
        sp.set(qn('w:before'), '0')
        sp.set(qn('w:after'),  str(after))
        pPr.append(sp)
        r = p.add_run(text)
        r.bold       = bold
        r.font.name  = FONT
        r.font.size  = Pt(size)
        return p

    para_centre(societe, bold=True, size=18, after=50)
    para_centre(f"{nom} {prenom}", bold=False, size=11, after=50)
    para_centre(f"Fiche horaire mensuelle — {MOIS_FR[mois]} {annee}",
                bold=True, size=11, after=100)

    jours = get_jours_mois(annee, mois)

    # Séparateurs après chaque dimanche
    lignes = []
    for j in jours:
        lignes.append(j)
        if j['dimanche']:
            lignes.append(None)   # séparateur
    # Supprimer dernier séparateur si en fin de liste
    if lignes and lignes[-1] is None:
        lignes.pop()

    table = doc.add_table(rows=0, cols=2)

    # ── En-tête ──
    hrow = table.add_row()
    set_row_height(hrow, 340)
    for i, txt in enumerate(('Date', 'Heures')):
        c = hrow.cells[i]
        set_shading(c, COLOR_HEADER)
        set_border(c)
        set_cell_margins(c, top=80, bottom=80, left=120, right=120)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(txt)
        r.bold = True
        r.font.name  = FONT
        r.font.size  = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # ── Jours ──
    for entry in lignes:
        if entry is None:
            sr = table.add_row()
            set_row_height(sr, 80)
            for c in sr.cells:
                set_shading(c, COLOR_WHITE)
                set_no_border(c)
            continue

        row = table.add_row()
        set_row_height(row, 255)

        avant_entree = (date_entree is not None
                        and entry['date'] < date_entree)

        fill = COLOR_WHITE
        if entry['ferie']:
            fill = COLOR_FERIE
        elif entry['dimanche']:
            fill = COLOR_DIM

        # Colonne Date
        cd = row.cells[0]
        set_shading(cd, fill)
        set_border(cd)
        set_cell_margins(cd, top=60, bottom=60, left=120, right=120)
        p  = cd.paragraphs[0]
        r  = p.add_run(entry['label'])
        r.font.name = FONT
        r.font.size = Pt(9)
        r.italic    = entry['weekend']

        # Colonne Heures
        ch = row.cells[1]
        set_shading(ch, COLOR_WHITE)
        set_border(ch)
        set_cell_margins(ch, top=60, bottom=60, left=80, right=80)
        if avant_entree:
            p2 = ch.paragraphs[0]
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = p2.add_run('X')
            r2.bold = True
            r2.font.name = FONT
            r2.font.size = Pt(9)

    # ── TOTAL ──
    tr = table.add_row()
    set_row_height(tr, 340)
    ct = tr.cells[0]
    set_shading(ct, COLOR_HEADER)
    set_border(ct)
    set_cell_margins(ct, top=80, bottom=80, left=120, right=120)
    p = ct.paragraphs[0]
    r = p.add_run('TOTAL')
    r.bold = True
    r.font.name  = FONT
    r.font.size  = Pt(9)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    ch2 = tr.cells[1]
    set_shading(ch2, COLOR_WHITE)
    set_border(ch2)

    fix_table_widths(table)

    # Signature
    ps = doc.add_paragraph()
    ps.paragraph_format.space_before = Pt(10)
    ps.add_run('Signature employé :').font.name = FONT

# ─── Interface Streamlit ──────────────────────────────────────────────────────
st.set_page_config(page_title="Fiches Horaires", page_icon="📋", layout="centered")
st.title("📋 Générateur de Fiches Horaires")

SOCIETES = ['GEA GAUMETOUS', 'Fed Agri', 'Feduline', 'Feduline II',
            'Paulhiac', 'Les Jardins de Casseneuil']
# Génère dynamiquement les 24 prochains mois + les 3 mois passés
def _gen_mois_options():
    today = datetime.date.today()
    start = datetime.date(today.year, today.month, 1)
    # 3 mois en arrière
    opts = {}
    for delta in range(-3, 25):
        y = start.year + (start.month - 1 + delta) // 12
        m = (start.month - 1 + delta) % 12 + 1
        opts[f"{MOIS_FR[m]} {y}"] = (y, m)
    return opts
MOIS_OPTIONS = _gen_mois_options()

tab1, tab2 = st.tabs(["✏️ Saisie manuelle", "📂 Import liste Excel"])

# ── Onglet 1 ──────────────────────────────────────────────────────────────────
with tab1:
    st.subheader("Ajouter des salariés")
    if 'salaries' not in st.session_state:
        st.session_state.salaries = []

    c1, c2, c3 = st.columns([2, 2, 2])
    with c1:
        nom_i    = st.text_input("Nom de famille", key="nom").upper()
    with c2:
        prenom_i = st.text_input("Prénom", key="prenom")
    with c3:
        soc_i    = st.selectbox("Société", SOCIETES, key="soc")

    c4, c5 = st.columns(2)
    with c4:
        mois_i = st.selectbox("Mois", list(MOIS_OPTIONS.keys()), key="mois",
                               index=list(MOIS_OPTIONS.keys()).index("Juillet 2026"))
    with c5:
        entree_i = st.date_input("Date d'entrée (optionnel)", value=None, key="entree")

    if st.button("➕ Ajouter à la liste"):
        if nom_i and prenom_i:
            st.session_state.salaries.append({
                'nom': nom_i, 'prenom': prenom_i, 'societe': soc_i,
                'mois': mois_i,
                'date_entree': str(entree_i) if entree_i else None,
            })
            st.success(f"{nom_i} {prenom_i} ajouté ✓")
        else:
            st.warning("Nom et prénom requis.")

    if st.session_state.salaries:
        st.divider()
        st.subheader(f"Liste ({len(st.session_state.salaries)} salarié(s))")
        st.dataframe(pd.DataFrame(st.session_state.salaries),
                     use_container_width=True, hide_index=True)

        ca, cb = st.columns(2)
        with ca:
            if st.button("🗑️ Vider la liste"):
                st.session_state.salaries = []
                st.rerun()
        with cb:
            if st.button("📄 Générer les fiches Word", type="primary"):
                doc = Document()
                for sec in doc.sections:
                    sec.top_margin    = Cm(1.2)
                    sec.bottom_margin = Cm(1.2)
                    sec.left_margin   = Cm(2)
                    sec.right_margin  = Cm(2)
                for i, s in enumerate(st.session_state.salaries):
                    annee, mois_num = MOIS_OPTIONS[s['mois']]
                    de = (datetime.date.fromisoformat(s['date_entree'])
                          if s['date_entree'] else None)
                    generer_fiche(doc, s['societe'], s['nom'], s['prenom'],
                                  annee, mois_num, de, premier=(i == 0))
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                st.download_button(
                    "⬇️ Télécharger le fichier Word", data=buf,
                    file_name="fiches_horaires.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ── Onglet 2 ──────────────────────────────────────────────────────────────────
with tab2:
    st.subheader("Importer une liste depuis Excel")
    st.info("Colonnes : **Nom**, **Prénom**, **Société**, **Mois** (ex: Juillet 2026), **Date entrée** (optionnel, JJ/MM/AAAA)")

    tpl = pd.DataFrame({
        'Nom': ['DUPONT', 'MARTIN'], 'Prénom': ['Jean', 'Marie'],
        'Société': ['GEA GAUMETOUS', 'Fed Agri'],
        'Mois': ['Juillet 2026', 'Juillet 2026'],
        'Date entrée': ['15/07/2026', ''],
    })
    bt = io.BytesIO()
    tpl.to_excel(bt, index=False)
    bt.seek(0)
    st.download_button("⬇️ Télécharger le modèle Excel", bt,
                       file_name="modele_salaries.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    up = st.file_uploader("Charger votre fichier Excel", type=['xlsx', 'xls'])
    if up:
        try:
            df = pd.read_excel(up)
            st.dataframe(df, use_container_width=True, hide_index=True)
            if st.button("📄 Générer toutes les fiches", type="primary"):
                doc = Document()
                for sec in doc.sections:
                    sec.top_margin    = Cm(1.2)
                    sec.bottom_margin = Cm(1.2)
                    sec.left_margin   = Cm(2)
                    sec.right_margin  = Cm(2)
                errors, count = [], 0
                for i, row in df.iterrows():
                    try:
                        nom    = str(row.get('Nom', '')).strip().upper()
                        prenom = str(row.get('Prénom', '')).strip()
                        soc    = str(row.get('Société', 'GEA GAUMETOUS')).strip()
                        mois_s = str(row.get('Mois', 'Juillet 2026')).strip()
                        der    = row.get('Date entrée', None)
                        if not nom or nom == 'NAN' or not prenom or prenom == 'nan':
                            continue
                        if mois_s not in MOIS_OPTIONS:
                            errors.append(f"Mois invalide : {nom} {prenom} → {mois_s}")
                            continue
                        annee, mois_num = MOIS_OPTIONS[mois_s]
                        de = None
                        if der and str(der) not in ('nan', 'NaT', ''):
                            try:
                                de = (der.date() if isinstance(der, datetime.datetime)
                                      else datetime.datetime.strptime(str(der), '%d/%m/%Y').date())
                            except:
                                pass
                        generer_fiche(doc, soc, nom, prenom, annee, mois_num,
                                      de, premier=(count == 0))
                        count += 1
                    except Exception as e:
                        errors.append(f"Erreur ligne {i+2} : {e}")
                for err in errors:
                    st.warning(err)
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                st.download_button(
                    f"⬇️ Télécharger ({count} fiches)", data=buf,
                    file_name="fiches_horaires.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Erreur lecture fichier : {e}")
