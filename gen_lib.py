import openpyxl, re, datetime, unicodedata
from difflib import SequenceMatcher
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart.data_source import StrRef, AxDataSource

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


def get_week_and_employees(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Preparation']
    periode = ws.cell(row=5, column=3).value  # "Du DD/MM/YYYY Au DD/MM/YYYY"
    m = re.search(r'Du (\d{2})/(\d{2})/(\d{4})', periode)
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    dt = datetime.date(y, mo, d)
    iso = dt.isocalendar()
    week = iso[1]

    def _to_float(v):
        try:
            return float(v) if v is not None else None
        except (TypeError, ValueError):
            return None

    row = 25
    employees = []  # (name, prep_row, articles_row, articles_count, commandes_count)
    while row <= ws.max_row:
        name = ws.cell(row=row, column=2).value
        if isinstance(name, str) and '(' in name and ')' in name:
            prep_row = row + 1
            articles_row = row + 3
            articles_count = _to_float(ws.cell(row=articles_row, column=3).value)
            commandes_count = _to_float(ws.cell(row=prep_row, column=3).value)
            employees.append((name.strip(), prep_row, articles_row, articles_count, commandes_count))
            row += 12
        else:
            row += 1
    return week, employees


# ---------------------------------------------------------------------------
# Extraction des heures "Drive" depuis un planning PDF (format Boulpat/Drive/Bazar)
# ---------------------------------------------------------------------------

_NAME_RE = re.compile(r'^[A-ZÀ-Ÿ\-]+$')
_HOUR_RE = re.compile(r'^\d{1,2}h\d{2}$')
_REPOS_TOKENS = {"REPOS", "MALADIE", "ACCIDENT", "CP"}
_NOISE_TOKENS = {
    "Pause", "min", "Reconnais", "SIGNATURE", "avoir", "effectué", "les", "horaires",
    "indiqués", "ci-dessus", "pris", "pauses", "indiquées", "sur", "chaque", "semaine",
    ":", "TOTAL", "Nb", "Heure",
}


def _cluster_rows(words, tol=2.5):
    words = sorted(words, key=lambda w: w['top'])
    rows, cur, cur_top = [], [], None
    for w in words:
        if cur_top is None or abs(w['top'] - cur_top) <= tol:
            cur.append(w)
            cur_top = w['top'] if cur_top is None else cur_top
        else:
            rows.append(cur)
            cur, cur_top = [w], w['top']
    if cur:
        rows.append(cur)
    for r in rows:
        r.sort(key=lambda w: w['x0'])
    return rows


def _hour_to_decimal(s):
    m = re.match(r'^(\d{1,2})h(\d{2})$', s)
    h, mi = int(m.group(1)), int(m.group(2))
    return h + mi / 60.0


def _group_tags(tags_row, gap_threshold=8.0):
    toks = sorted([w for w in tags_row if w['text'] not in _NOISE_TOKENS], key=lambda w: w['x0'])
    groups = []
    for w in toks:
        if groups and (w['x0'] - groups[-1]['x1']) <= gap_threshold:
            groups[-1]['text'] += ' ' + w['text']
            groups[-1]['x1'] = w['x1']
        else:
            groups.append({'text': w['text'], 'x0': w['x0'], 'x1': w['x1']})
    return groups


def parse_planning_pdf(pdf_path, department='DRIVE'):
    """Lit un planning hebdomadaire (format Boulpat/Drive/Bazar) et retourne un
    dict {nom_planning: heures_decimales} ne comptabilisant que les heures
    dont le rayon affecté correspond exactement a `department` (ex: DRIVE).
    Les employes absents toute la semaine (maladie/accident) ou n'ayant
    jamais travaille sur ce rayon obtiennent 0."""
    if pdfplumber is None:
        raise RuntimeError("pdfplumber n'est pas installe: impossible de lire le planning PDF.")

    blocks = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
            rows = _cluster_rows(words)
            name_rows_idx = [
                i for i, r in enumerate(rows)
                if r[0]['x0'] < 135 and _NAME_RE.match(r[0]['text']) and len(r[0]['text']) >= 2 and r[0]['top'] > 85
            ]
            for bi, idx in enumerate(name_rows_idx):
                end_idx = name_rows_idx[bi + 1] if bi + 1 < len(name_rows_idx) else len(rows)
                block_rows = rows[idx:end_idx]
                name_tokens = [w['text'] for w in block_rows[0] if w['x0'] < 135]
                blocks.append((' '.join(name_tokens), block_rows))

    result = {}
    for name, block_rows in blocks:
        hours_row, tags_row = None, None
        for r in block_rows:
            toks = [w['text'] for w in r]
            if toks and all(_HOUR_RE.match(t) for t in toks) and r[0]['x0'] >= 135:
                hours_row = r
            elif toks and all(t in _REPOS_TOKENS for t in toks):
                pass
            elif toks and all(_NAME_RE.match(t) for t in toks) and r[0]['x0'] >= 135:
                tags_row = r

        if not hours_row:
            result[name] = 0.0
            continue

        hour_tokens = sorted(hours_row, key=lambda w: w['x0'])
        groups = _group_tags(tags_row) if tags_row else []

        dept_hours = 0.0
        for i, htok in enumerate(hour_tokens):
            tag = groups[i]['text'].strip().upper() if i < len(groups) else None
            if tag == department.upper():
                dept_hours += _hour_to_decimal(htok['text'])
        result[name] = round(dept_hours, 2)

    return result


def _normalize_name_tokens(name):
    name = re.sub(r'\([^)]*\)', '', name)  # retire le matricule entre parentheses
    name = unicodedata.normalize('NFKD', name).encode('ascii', 'ignore').decode('ascii')
    name = re.sub(r'[^A-Za-z\s\-]', ' ', name)
    return [t for t in name.upper().replace('-', ' ').split() if t]


def _name_match_score(name_a, name_b):
    a_tokens = _normalize_name_tokens(name_a)
    b_tokens = _normalize_name_tokens(name_b)
    if not a_tokens or not b_tokens:
        return 0.0
    short, long_ = (a_tokens, b_tokens) if len(a_tokens) <= len(b_tokens) else (b_tokens, a_tokens)
    used, scores = set(), []
    for t in short:
        best, best_i = 0.0, None
        for i, u in enumerate(long_):
            if i in used:
                continue
            s = SequenceMatcher(None, t, u).ratio()
            if s > best:
                best, best_i = s, i
        if best_i is not None:
            used.add(best_i)
        scores.append(best)
    return sum(scores) / len(scores)


def extract_matricule(name):
    m = re.search(r'\(([^)]+)\)', name)
    return m.group(1).strip() if m else None


def match_planning_hours(planning_hours, employee_names, threshold=0.82):
    """Associe chaque employe (nom tel qu'il figure dans la Preparation) aux
    heures Drive du planning, en tolerant fautes de frappe et ordre nom/prenom
    inverse. Ne renvoie une valeur que pour les employes matches avec un score
    suffisant ; les autres sont absents du dict retourne (fallback manuel)."""
    matched = {}
    for emp_name in employee_names:
        best_score, best_hours = 0.0, None
        for plan_name, hours in planning_hours.items():
            score = _name_match_score(plan_name, emp_name)
            if score > best_score:
                best_score, best_hours = score, hours
        if best_score >= threshold:
            matched[emp_name] = best_hours
    return matched


def build(path, outpath, taux_actuelle=None, taux_precedente=None, planning_path=None, productivite_s1=None):
    """productivite_s1: dict {matricule: productivite_h_decimale} issu de la
    semaine precedente, utilise pour auto-remplir la colonne H. Retourne
    (semaine, nb_employes, productivite_calculee) ou productivite_calculee
    est un dict {matricule: productivite_h_decimale} pour cette semaine,
    a persister pour servir de S-1 la semaine suivante."""
    week, employees = get_week_and_employees(path)

    hours_map = {}
    if planning_path:
        planning_hours = parse_planning_pdf(planning_path)
        hours_map = match_planning_hours(planning_hours, [e[0] for e in employees])

    productivite_s1 = productivite_s1 or {}
    employee_productivity_this_week = {}

    # Pre-clean: round-trip once through openpyxl to strip Numbers-export style
    # cruft that otherwise makes LibreOffice hang when a chart is added later.
    clean_path = path.replace('.xlsx', '_clean.xlsx')
    wb_clean = openpyxl.load_workbook(path, data_only=False)
    wb_clean.save(clean_path)
    wb = openpyxl.load_workbook(clean_path, data_only=False)

    prep_sheet_name = f'Preparation Semaine {week}'
    prod_sheet_name = f'Feuil2 Productivité Semaine {week}'
    assert len(prod_sheet_name) <= 31, prod_sheet_name

    ws_prep = wb['Preparation']
    ws_prep.title = prep_sheet_name

    ws = wb['Feuil2']
    ws.title = prod_sheet_name

    arial = 'Arial'
    n = len(employees)
    first_data_row = 3
    last_data_row = first_data_row + n - 1

    # Title
    ws['A1'] = 'Tableau 2 - Productivité'
    ws['A1'].font = Font(name=arial, bold=True, size=14)

    headers = ['Employé', 'Heure travaillée', "Nbr d'articles préparés",
               'Nbr de commande préparés', 'Productivité /H', 'Productivité minutes',
               'Nombre de commande à l\'heure']
    for i, h in enumerate(headers):
        c = ws.cell(row=2, column=1+i, value=h)
        c.font = Font(name=arial, bold=True)
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    blue_font = Font(name=arial, color='0000FF')

    for i, (name, prep_row, art_row, articles_count, commandes_count) in enumerate(employees):
        r = first_data_row + i
        matricule = extract_matricule(name)
        ws.cell(row=r, column=1, value=name).font = Font(name=arial)

        b_value = None
        cb = ws.cell(row=r, column=2)
        if name in hours_map:
            b_value = hours_map[name]
            cb.value = b_value
            cb.fill = green_fill
            cb.font = Font(name=arial)
        else:
            cb.value = None
            cb.fill = yellow_fill
            cb.font = blue_font
        cb.number_format = '0.00'

        # C/D/E/F/G : quand les heures (B) sont deja connues (auto-remplies
        # depuis le planning), on ecrit des valeurs litterales calculees en
        # Python plutot que des formules — certains lecteurs (Numbers, aperçus)
        # n'executent pas toujours le recalcul des formules a l'ouverture, ce
        # qui laissait ces colonnes a 0/vide. Quand B reste manuel (jaune), on
        # garde les formules d'origine pour un calcul live des que l'utilisateur
        # saisit une valeur dans Excel.
        e_value = None
        if b_value is not None:
            cc = ws.cell(row=r, column=3, value=articles_count if articles_count is not None else 0)
            cd = ws.cell(row=r, column=4, value=commandes_count if commandes_count is not None else 0)
            e_value = (articles_count / b_value) if (articles_count is not None and b_value) else None
            ce = ws.cell(row=r, column=5, value=e_value)
            cf = ws.cell(row=r, column=6, value=(e_value / 60) if e_value is not None else None)
            cg = ws.cell(row=r, column=7, value=(commandes_count / b_value) if (commandes_count is not None and b_value) else None)
        else:
            cc = ws.cell(row=r, column=3, value=f"='{prep_sheet_name}'!C{art_row}")
            cd = ws.cell(row=r, column=4, value=f"='{prep_sheet_name}'!C{prep_row}")
            ce = ws.cell(row=r, column=5, value=f"=IFERROR(C{r}/B{r},\"\")")
            cf = ws.cell(row=r, column=6, value=f"=IFERROR(E{r}/60,\"\")")
            cg = ws.cell(row=r, column=7, value=f"=IFERROR(D{r}/B{r},\"\")")
        cc.font = Font(name=arial, color='008000')
        cd.font = Font(name=arial, color='008000')
        ce.font = Font(name=arial)
        ce.number_format = '0.00'
        cf.font = Font(name=arial)
        cf.number_format = '0.00'
        cg.font = Font(name=arial)
        cg.number_format = '0.00'

        h_value = None
        ch = ws.cell(row=r, column=8)
        if matricule and matricule in productivite_s1:
            h_value = productivite_s1[matricule]
            ch.value = h_value
            ch.fill = green_fill
            ch.font = Font(name=arial)
        else:
            ch.value = None
            ch.fill = yellow_fill
            ch.font = blue_font
        ch.number_format = '0.00'

        # % Evolution vs S-1 : difference entre Productivite/H (E, semaine
        # actuelle) et Productivite/H S-1 (H), exprimee en % de la valeur
        # ACTUELLE (E) — donc (E-H)/E*100, pas /H.
        ci = ws.cell(row=r, column=9)
        if e_value is not None and h_value is not None and e_value != 0:
            arrow = '▲ ' if e_value >= h_value else '▼ '
            pct = round(abs(e_value - h_value) / e_value * 100, 1)
            ci.value = f"{arrow}{pct}%"
        elif b_value is not None and h_value is not None and e_value == 0:
            # Division par zero (productivite actuelle = 0) : rien de comparable.
            ci.value = ""
        elif b_value is not None and h_value is None:
            # B connu mais S-1 pas encore disponible : rien a comparer.
            ci.value = ""
        else:
            ci.value = (f"=IFERROR(IF(OR(E{r}=\"\",H{r}=\"\"),\"\",IF(E{r}>=H{r},\"▲ \",\"▼ \")"
                        f"&ROUND(ABS(E{r}-H{r})/E{r}*100,1)&\"%\"),\"\")")
        ci.font = Font(name=arial, bold=True)
        ci.alignment = Alignment(horizontal='center')

        # Productivite calculee cette semaine (pour servir de S-1 la semaine prochaine)
        if matricule and b_value and articles_count is not None:
            employee_productivity_this_week[matricule] = {
                'nom': name,
                'productivite_h': round(articles_count / b_value, 2),
            }

    headers2 = ['Productivité /H S-1', '% Évolution vs S-1']
    for i, h in enumerate(headers2):
        c = ws.cell(row=2, column=8+i, value=h)
        c.font = Font(name=arial, bold=True)
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

    ws.conditional_formatting.add(
        f'I{first_data_row}:I{last_data_row}',
        FormulaRule(formula=[f'AND($E{first_data_row}<>\"\",$H{first_data_row}<>\"\",$E{first_data_row}>=$H{first_data_row})'],
                    font=Font(name=arial, color='008000', bold=True)))
    ws.conditional_formatting.add(
        f'I{first_data_row}:I{last_data_row}',
        FormulaRule(formula=[f'AND($E{first_data_row}<>\"\",$H{first_data_row}<>\"\",$E{first_data_row}<$H{first_data_row})'],
                    font=Font(name=arial, color='FF0000', bold=True)))

    widths = {'A': 32, 'B': 16, 'C': 20, 'D': 20, 'E': 16, 'F': 16, 'G': 20, 'H': 18, 'I': 18}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    k_headers = ['Rang', 'Employé (classé)', 'Productivité /H (classée)']
    for i, h in enumerate(k_headers):
        c = ws.cell(row=2, column=11+i, value=h)
        c.font = Font(name=arial, bold=True)
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

    e_range = f"$E${first_data_row}:$E${last_data_row}"
    a_range = f"$A${first_data_row}:$A${last_data_row}"

    for i in range(n):
        r = first_data_row + i
        ck = ws.cell(row=r, column=11, value=f"=ROW()-{first_data_row-1}")
        ck.font = Font(name=arial)
        cm = ws.cell(row=r, column=13, value=f"=IFERROR(LARGE({e_range},K{r}),\"\")")
        cm.font = Font(name=arial)
        cm.number_format = '0.00'
        cl = ws.cell(row=r, column=12,
                      value=f"=IF(M{r}=\"\",\"\",IFERROR(INDEX({a_range},MATCH(M{r},{e_range},0)),\"\"))")
        cl.font = Font(name=arial)

    widths2 = {'K': 8, 'L': 32, 'M': 22}
    for col, w in widths2.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells('O1:P1')
    ct = ws['O1']
    ct.value = 'Indicateur - Taux de rupture (avant substitution)'
    ct.font = Font(name=arial, bold=True)
    ct.alignment = Alignment(horizontal='center', wrap_text=True)

    ws['O2'] = 'Semaine précédente (S-1)'
    ws['O2'].font = Font(name=arial)
    p2 = ws['P2']
    p2.fill = yellow_fill
    p2.font = blue_font
    p2.number_format = '0.0%'

    ws['O3'] = 'Semaine actuelle'
    ws['O3'].font = Font(name=arial)
    p3 = ws['P3']
    p3.fill = yellow_fill
    p3.font = blue_font
    p3.number_format = '0.0%'
    if taux_actuelle is not None:
        p3.value = taux_actuelle
    if taux_precedente is not None:
        p2.value = taux_precedente

    ws['O4'] = 'Évolution'
    ws['O4'].font = Font(name=arial, bold=True)
    p4 = ws['P4']
    p4.value = ('=IFERROR(IF(OR(P2="",P3=""),"",ROUND(P2*100,1)&"%"&"  "'
                '&IF(P3<=P2,"▼","▲")&"  "&ROUND(P3*100,1)&"%"),"")')
    p4.font = Font(name=arial, bold=True)
    p4.alignment = Alignment(horizontal='center')

    ws.conditional_formatting.add(
        'P4', FormulaRule(formula=['AND(P2<>"",P3<>"",P3<=P2)'],
                           font=Font(name=arial, color='008000', bold=True, italic=True)))
    ws.conditional_formatting.add(
        'P4', FormulaRule(formula=['AND(P2<>"",P3<>"",P3>P2)'],
                           font=Font(name=arial, color='FF0000', bold=True, italic=True)))

    widths3 = {'O': 24, 'P': 20}
    for col, w in widths3.items():
        ws.column_dimensions[col].width = w

    wb.calculation.fullCalcOnLoad = True
    wb.save(outpath)
    wb = openpyxl.load_workbook(outpath)
    ws = wb[prod_sheet_name]

    chart = BarChart()
    chart.type = 'col'
    chart.title = 'Productivité à l\'heure par employé (du plus grand au plus petit)'
    chart.y_axis.title = 'Productivité /H'
    chart.x_axis.title = 'Employé'
    chart.style = 10
    chart.width = max(24, n * 2.2)
    chart.height = 12

    data = Reference(ws, min_col=13, min_row=2, max_row=last_data_row)
    cats = Reference(ws, min_col=12, min_row=first_data_row, max_row=last_data_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    cat_str_ref = StrRef(f=str(cats))
    chart.series[0].cat = AxDataSource(strRef=cat_str_ref)
    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showLegendKey = False
    chart.dLbls.showCatName = False
    chart.dLbls.showSerName = False
    chart.legend = None

    body_pr = RichTextProperties(rot=-2700000, vert='horz')
    cp = CharacterProperties(sz=900)
    pp = ParagraphProperties(defRPr=cp)
    rich = RichText(bodyPr=body_pr, p=[Paragraph(pPr=pp, endParaRPr=cp)])
    chart.x_axis.txPr = rich
    chart.x_axis.delete = False
    chart.y_axis.delete = False

    chart_anchor_row = last_data_row + 4
    ws.add_chart(chart, f'A{chart_anchor_row}')

    wb.calculation.fullCalcOnLoad = True
    wb.save(outpath)
    return week, n, employee_productivity_this_week
