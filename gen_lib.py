import openpyxl, re, datetime, unicodedata
from difflib import SequenceMatcher
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart.data_source import StrRef, AxDataSource
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as _rl_canvas
    from reportlab.lib import colors as _rl_colors
    from reportlab.pdfbase.pdfmetrics import stringWidth as _rl_stringWidth
except ImportError:
    _rl_canvas = None

# Matricules exclus du classement/podium car responsables (pas des collaborateurs
# a classer). A adapter si l'organisation change.
EXCLUDED_FROM_RANKING = {
    "MONTESCOT3",   # Adrien Navaro - responsable
    "PR0911201",    # Maelle Gendre - responsable
}


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


def get_periode(path):
    """Renvoie (date_debut, date_fin) de la periode reellement couverte par
    le fichier Preparation (cellule "Du DD/MM/YYYY Au DD/MM/YYYY"). Renvoie
    (None, None) si le format n'est pas reconnu."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Preparation']
    periode = ws.cell(row=5, column=3).value
    m = re.search(r'Du (\d{2})/(\d{2})/(\d{4}) [Aa]u (\d{2})/(\d{2})/(\d{4})', periode or '')
    if not m:
        return None, None
    d1, mo1, y1, d2, mo2, y2 = (int(g) for g in m.groups())
    return datetime.date(y1, mo1, d1), datetime.date(y2, mo2, d2)


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


_WEEKDAYS = {"Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"}
_DATE_DM_RE = re.compile(r'^(\d{2})/(\d{2})$')


def _extract_day_columns(page):
    """Retourne une liste [(x0_debut_colonne, date), ...] triee par x0, une
    entree par jour de la semaine affiche sur cette page du planning (en se
    basant sur les colonnes "Lundi 29/06", "Mardi 30/06", ... et le sous-en-tete
    "Matin" qui marque le debut de chaque colonne jour). Retourne [] si les
    en-tetes attendus ne sont pas trouves (le filtrage par date est alors
    simplement desactive, sans erreur)."""
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    header_words = [w for w in words if w['top'] < 90]
    if not header_words:
        return []

    header_text = ' '.join(w['text'] for w in sorted(header_words, key=lambda w: (w['top'], w['x0'])))
    year_m = re.search(r'(\d{2})/(\d{2})/(\d{4})', header_text)
    if not year_m:
        return []
    year = int(year_m.group(3))

    day_row = sorted([w for w in words if 60 <= w['top'] <= 75], key=lambda w: w['x0'])
    dates = []
    i = 0
    while i < len(day_row):
        w = day_row[i]
        if w['text'] in _WEEKDAYS and i + 1 < len(day_row):
            dm = _DATE_DM_RE.match(day_row[i + 1]['text'])
            if dm:
                try:
                    dates.append(datetime.date(year, int(dm.group(2)), int(dm.group(1))))
                except ValueError:
                    dates.append(None)
                i += 2
                continue
        i += 1

    matin_row = sorted(
        [w for w in words if 76 <= w['top'] <= 85 and w['text'] == 'Matin'],
        key=lambda w: w['x0'])
    starts = [w['x0'] for w in matin_row]

    if not starts or len(starts) != len(dates) or any(d is None for d in dates):
        return []
    return list(zip(starts, dates))


def _date_for_x(x0, day_columns):
    """Renvoie la date de la colonne jour dans laquelle x0 tombe, a partir
    d'une liste [(x0_debut, date), ...] triee par x0 croissant."""
    chosen = None
    for start_x, d in day_columns:
        if x0 >= start_x - 3:
            chosen = d
        else:
            break
    return chosen


def parse_planning_pdf(pdf_path, department='DRIVE', date_start=None, date_end=None):
    """Lit un planning hebdomadaire (format Boulpat/Drive/Bazar) et retourne un
    dict {nom_planning: heures_decimales} ne comptabilisant que les heures
    dont le rayon affecté correspond exactement a `department` (ex: DRIVE).
    Les employes absents toute la semaine (maladie/accident) ou n'ayant
    jamais travaille sur ce rayon obtiennent 0.

    Si date_start/date_end sont fournis (periode reelle du fichier
    Preparation), seules les heures dont la colonne jour du planning tombe
    dans cette periode sont comptabilisees — utile quand le planning couvre
    une semaine calendaire complete (ex: Lundi a Dimanche) mais que le suivi
    de productivite ne porte que sur une partie de cette semaine (ex:
    Mercredi a Dimanche pour la premiere semaine d'un nouveau cycle)."""
    if pdfplumber is None:
        raise RuntimeError("pdfplumber n'est pas installe: impossible de lire le planning PDF.")

    result = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
            rows = _cluster_rows(words)
            day_columns = _extract_day_columns(page) if (date_start or date_end) else []

            name_rows_idx = [
                i for i, r in enumerate(rows)
                if r[0]['x0'] < 135 and _NAME_RE.match(r[0]['text']) and len(r[0]['text']) >= 2 and r[0]['top'] > 85
            ]
            for bi, idx in enumerate(name_rows_idx):
                end_idx = name_rows_idx[bi + 1] if bi + 1 < len(name_rows_idx) else len(rows)
                block_rows = rows[idx:end_idx]
                name_tokens = [w['text'] for w in block_rows[0] if w['x0'] < 135]
                name = ' '.join(name_tokens)

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
                    if tag != department.upper():
                        continue
                    if day_columns:
                        d = _date_for_x(htok['x0'], day_columns)
                        if d is not None and date_start is not None and d < date_start:
                            continue
                        if d is not None and date_end is not None and d > date_end:
                            continue
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
    periode_debut, periode_fin = get_periode(path)

    hours_map = {}
    if planning_path:
        # Ne compter que les heures DRIVE dont la date tombe dans la periode
        # reellement couverte par le fichier Preparation : le planning peut
        # afficher une semaine calendaire complete (Lundi-Dimanche) alors que
        # le suivi de productivite ne porte que sur une partie de celle-ci
        # (ex: premiere semaine d'un cycle qui demarre un mercredi).
        planning_hours = parse_planning_pdf(planning_path, date_start=periode_debut, date_end=periode_fin)
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
        evo_pct, evo_up = None, None
        ci = ws.cell(row=r, column=9)
        if e_value is not None and h_value is not None and e_value != 0:
            evo_up = e_value >= h_value
            evo_pct = round(abs(e_value - h_value) / e_value * 100, 1)
            arrow = '▲ ' if evo_up else '▼ '
            ci.value = f"{arrow}{evo_pct}%"
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
            entry = {
                'nom': name,
                'productivite_h': round(articles_count / b_value, 2),
            }
            if evo_pct is not None:
                entry['evolution_pct'] = evo_pct
                entry['evolution_up'] = evo_up
            employee_productivity_this_week[matricule] = entry

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

    # Une couleur distincte par employe (barre), plutot qu'une seule couleur
    # pour toute la serie.
    palette = [
        '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47',
        '264478', '9E480E', '636363', '997300', '255E91', '43682B',
        'C00000', '7030A0', '00B0F0', 'FF66CC', '00B050', 'BF8F00',
        '203864', '833C00',
    ]
    chart.series[0].graphicalProperties.varyColors = True
    data_points = []
    for idx in range(n):
        color = palette[idx % len(palette)]
        dp = DataPoint(idx=idx)
        dp.graphicalProperties = GraphicalProperties(solidFill=color)
        data_points.append(dp)
    chart.series[0].data_points = data_points

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


# ---------------------------------------------------------------------------
# PDF "Podium" - affichage visuel hebdomadaire (charte Intermarché) a afficher
# en salle de pause. Design valide avec Adrien le 02/08/2026.
# ---------------------------------------------------------------------------

_PDM_RED = None
_PDM_BLACK = None


def _pdm_colors():
    global _PDM_RED, _PDM_BLACK
    if _PDM_RED is None:
        _PDM_RED = _rl_colors.HexColor('#E2001A')
        _PDM_BLACK = _rl_colors.HexColor('#1A1A1A')
    return _PDM_RED, _PDM_BLACK


def _pdm_rounded_card(c, x, y, w, h, fill, stroke=None, radius=10, stroke_width=1.2):
    c.saveState()
    c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke)
        c.setLineWidth(stroke_width)
        c.roundRect(x, y, w, h, radius, fill=1, stroke=1)
    else:
        c.roundRect(x, y, w, h, radius, fill=1, stroke=0)
    c.restoreState()


def _pdm_center_text(c, text, cx, y, font, size, color):
    c.setFont(font, size)
    c.setFillColor(color)
    c.drawCentredString(cx, y, text)


def _pdm_evolution_pill(c, cx, y, up, pct, WHITE, GREEN, RED_NEG, GREY):
    if pct is None:
        label, color = "NOUVEAU", GREY
    else:
        label = f"{'▲' if up else '▼'} {pct:.1f}%"
        color = GREEN if up else RED_NEG
    c.setFont("Helvetica-Bold", 10)
    tw = _rl_stringWidth(label, "Helvetica-Bold", 10)
    pad = 10
    pw, ph = tw + 2 * pad, 18
    _pdm_rounded_card(c, cx - pw / 2, y, pw, ph, color, radius=9)
    c.setFillColor(WHITE)
    c.drawCentredString(cx, y + 5, label)


def _pdm_medal_circle(c, cx, cy, r, color, text, WHITE):
    c.setFillColor(color)
    c.circle(cx, cy, r, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", r * 0.85)
    c.drawCentredString(cx, cy - r * 0.32, text)


def _pdm_logo_mark(c, x, y, size, RED, WHITE):
    """Badge enseigne : carre rouge arrondi + panier stylise (icone generique)."""
    _pdm_rounded_card(c, x, y, size, size, RED, radius=size * 0.22)
    c.saveState()
    cx, cy = x + size / 2, y + size / 2 - size * 0.04
    c.setStrokeColor(WHITE)
    c.setLineWidth(size * 0.055)
    c.arc(cx - size * 0.16, cy + size * 0.02, cx + size * 0.16, cy + size * 0.32, 0, 180)
    basket_w, basket_h = size * 0.46, size * 0.24
    bx, by = cx - basket_w / 2, cy - basket_h / 2
    p = c.beginPath()
    p.moveTo(bx + basket_w * 0.08, by + basket_h)
    p.lineTo(bx, by)
    p.lineTo(bx + basket_w, by)
    p.lineTo(bx + basket_w * 0.92, by + basket_h)
    p.close()
    c.setFillColor(WHITE)
    c.drawPath(p, fill=1, stroke=0)
    c.circle(bx + basket_w * 0.2, by - size * 0.05, size * 0.035, fill=1, stroke=0)
    c.circle(bx + basket_w * 0.8, by - size * 0.05, size * 0.035, fill=1, stroke=0)
    c.restoreState()


def _pdm_clean_name(nom):
    base = re.sub(r'\([^)]*\)', '', nom).strip()
    return base.title()


def build_ranking(employee_productivity, excluded_matricules=None):
    """Construit la liste classee (desc. par productivite_h) a partir du dict
    retourne par build(), en excluant les responsables."""
    excluded = excluded_matricules if excluded_matricules is not None else EXCLUDED_FROM_RANKING
    rows = []
    for matricule, data in (employee_productivity or {}).items():
        if matricule in excluded:
            continue
        if data.get('productivite_h') is None:
            continue
        rows.append({
            'matricule': matricule,
            'name': _pdm_clean_name(data.get('nom', matricule)),
            'prod': data['productivite_h'],
            'evo': data.get('evolution_pct'),
            'up': data.get('evolution_up'),
        })
    rows.sort(key=lambda r: r['prod'], reverse=True)
    return rows


def generate_podium_pdf(pdf_path, week, employee_productivity, taux_actuelle=None,
                         taux_precedente=None, enseigne="INTERMARCHÉ", magasin="MONTESCOT",
                         excluded_matricules=None):
    """Genere le PDF "podium" hebdomadaire (charte Intermarche) a partir des
    resultats de build(). Retourne le nombre de collaborateurs classes."""
    if _rl_canvas is None:
        raise RuntimeError("reportlab n'est pas installe: impossible de generer le PDF podium.")

    RED, BLACK = _pdm_colors()
    WHITE = _rl_colors.white
    GOLD = _rl_colors.HexColor('#D9A400')
    SILVER = _rl_colors.HexColor('#9AA0A6')
    BRONZE = _rl_colors.HexColor('#B5651D')
    GREEN = _rl_colors.HexColor('#1E7B34')
    RED_NEG = _rl_colors.HexColor('#C62828')
    GREY_BG = _rl_colors.HexColor('#F3F3F3')
    GREY_TXT = _rl_colors.HexColor('#5B5B5B')
    GREY_PILL = _rl_colors.HexColor('#9E9E9E')

    ranking = build_ranking(employee_productivity, excluded_matricules)
    W, H = A4
    c = _rl_canvas.Canvas(pdf_path, pagesize=A4)

    # ---- Header
    header_h = 108
    c.setFillColor(BLACK)
    c.rect(0, H - header_h, W, header_h, fill=1, stroke=0)
    c.setFillColor(RED)
    c.rect(0, H - header_h - 6, W, 6, fill=1, stroke=0)

    logo_size = 52
    logo_x, logo_y = 36, H - 34 - logo_size
    _pdm_logo_mark(c, logo_x, logo_y, logo_size, RED, WHITE)

    text_x = logo_x + logo_size + 14
    c.setFont("Helvetica-Bold", 25)
    c.setFillColor(WHITE)
    c.drawString(text_x, H - 46, enseigne)
    c.setFillColor(RED)
    tw = _rl_stringWidth(enseigne + " ", "Helvetica-Bold", 25)
    c.drawString(text_x + tw, H - 46, magasin)

    c.setFont("Helvetica-Bold", 11)
    band_label = "PERFORMANCE DRIVE"
    btw = _rl_stringWidth(band_label, "Helvetica-Bold", 11)
    bpad = 8
    _pdm_rounded_card(c, text_x, H - 74, btw + 2 * bpad, 17, RED, radius=8)
    c.setFillColor(WHITE)
    c.drawString(text_x + bpad, H - 70, band_label)
    c.setFont("Helvetica", 10)
    c.setFillColor(_rl_colors.HexColor('#CCCCCC'))
    c.drawString(text_x + btw + 2 * bpad + 10, H - 70, "Classement hebdomadaire des collaborateurs")

    pill_label = f"SEMAINE {week}"
    c.setFont("Helvetica-Bold", 13)
    ptw = _rl_stringWidth(pill_label, "Helvetica-Bold", 13)
    pw, ph = ptw + 28, 30
    px, py = W - 36 - pw, H - header_h / 2 - ph / 2
    _pdm_rounded_card(c, px, py, pw, ph, RED, radius=15)
    c.setFillColor(WHITE)
    c.drawCentredString(px + pw / 2, py + ph / 2 - 5, pill_label)

    # ---- Podium top 3
    podium_top = H - header_h - 18
    card_w, gap = 152, 14
    total_w = card_w * 3 + gap * 2
    start_x = (W - total_w) / 2
    order = [1, 0, 2]
    medal_colors = {0: GOLD, 1: SILVER, 2: BRONZE}
    card_h_map = {0: 172, 1: 148, 2: 148}
    bar_h_map = {0: 150, 1: 108, 2: 78}
    baseline = podium_top - 200

    top3 = ranking[:3]
    for slot, idx in enumerate(order):
        if idx >= len(top3):
            continue
        x = start_x + slot * (card_w + gap)
        p = top3[idx]
        card_h = card_h_map[idx]
        card_y = podium_top - card_h
        color = medal_colors[idx]

        _pdm_rounded_card(c, x, card_y, card_w, card_h, WHITE, stroke=color, stroke_width=2.2, radius=12)
        cx = x + card_w / 2
        _pdm_medal_circle(c, cx, card_y + card_h - 8, 22, color, str(idx + 1), WHITE)

        c.setFont("Helvetica-Bold", 11)
        name = p["name"]
        if _rl_stringWidth(name, "Helvetica-Bold", 11) > card_w - 16:
            parts = name.split(" ", 1)
            _pdm_center_text(c, parts[0], cx, card_y + card_h - 46, "Helvetica-Bold", 11, BLACK)
            _pdm_center_text(c, parts[1] if len(parts) > 1 else "", cx, card_y + card_h - 59, "Helvetica-Bold", 11, BLACK)
            name_bottom = card_y + card_h - 59
        else:
            _pdm_center_text(c, name, cx, card_y + card_h - 46, "Helvetica-Bold", 11, BLACK)
            name_bottom = card_y + card_h - 46

        _pdm_center_text(c, f"{p['prod']:.1f}", cx, name_bottom - 24, "Helvetica-Bold", 20, BLACK)
        _pdm_center_text(c, "articles / heure", cx, name_bottom - 36, "Helvetica", 7.5, GREY_TXT)
        _pdm_evolution_pill(c, cx, card_y + 10, p["up"], p["evo"], WHITE, GREEN, RED_NEG, GREY_PILL)

        bar_h = bar_h_map[idx]
        bar_y = baseline - bar_h
        c.setFillColor(color)
        c.rect(x, bar_y, card_w, bar_h, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 34)
        c.setFillColor(WHITE)
        c.drawCentredString(cx, bar_y + bar_h / 2 - 12, str(idx + 1))

    list_top = baseline - max(bar_h_map.values()) - 22

    # ---- Classement 4e et suivants
    row_h = 24
    header_row_h = 22
    list_rows = ranking[3:]
    y = list_top
    c.setFillColor(RED)
    c.rect(36, y - header_row_h, W - 72, header_row_h, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(50, y - header_row_h + 7, "RANG")
    c.drawString(100, y - header_row_h + 7, "COLLABORATEUR")
    c.drawString(360, y - header_row_h + 7, "PRODUCTIVITÉ /H")
    c.drawString(480, y - header_row_h + 7, "ÉVOLUTION S-1")
    y -= header_row_h

    for i, p in enumerate(list_rows):
        rank = i + 4
        row_y = y - row_h
        if i % 2 == 0:
            c.setFillColor(GREY_BG)
            c.rect(36, row_y, W - 72, row_h, fill=1, stroke=0)
        c.setFillColor(BLACK)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(50, row_y + 8, f"{rank}e")
        c.setFont("Helvetica", 10)
        c.drawString(100, row_y + 8, p["name"])
        c.setFont("Helvetica-Bold", 10)
        c.drawString(360, row_y + 8, f"{p['prod']:.1f} art./h")
        if p["evo"] is None:
            c.setFillColor(GREY_PILL)
            c.drawString(480, row_y + 8, "NOUVEAU")
        else:
            arrow = "▲" if p["up"] else "▼"
            c.setFillColor(GREEN if p["up"] else RED_NEG)
            c.drawString(480, row_y + 8, f"{arrow} {p['evo']:.1f}%")
        y = row_y

    list_bottom = y

    # ---- Taux de rupture
    taux_h = 108
    box_y = list_bottom - 24 - taux_h
    _pdm_rounded_card(c, 36, box_y, W - 72, taux_h, BLACK, radius=12)
    c.setFillColor(RED)
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(W / 2, box_y + taux_h - 24, "TAUX DE RUPTURE")

    half_w = (W - 72) / 2
    ts1 = taux_precedente if taux_precedente is not None else 0.0
    tac = taux_actuelle if taux_actuelle is not None else 0.0
    _pdm_center_text(c, "SEMAINE PRÉCÉDENTE (S-1)", 36 + half_w / 2, box_y + taux_h - 50, "Helvetica", 9, _rl_colors.HexColor('#BBBBBB'))
    _pdm_center_text(c, f"{ts1*100:.2f}%", 36 + half_w / 2, box_y + 22, "Helvetica-Bold", 26, WHITE)
    _pdm_center_text(c, "CETTE SEMAINE", 36 + half_w + half_w / 2, box_y + taux_h - 50, "Helvetica", 9, _rl_colors.HexColor('#BBBBBB'))
    delta = tac - ts1
    worse = delta > 0
    delta_color = RED_NEG if worse else GREEN
    _pdm_center_text(c, f"{tac*100:.2f}%", 36 + half_w + half_w / 2, box_y + 22, "Helvetica-Bold", 26, delta_color)

    c.setStrokeColor(_rl_colors.HexColor('#444444'))
    c.setLineWidth(1)
    c.line(36 + half_w, box_y + 14, 36 + half_w, box_y + taux_h - 40)

    delta_label = f"{'+' if worse else ''}{delta*100:.2f} pt {'▲' if worse else '▼'}"
    c.setFont("Helvetica-Bold", 10)
    dtw = _rl_stringWidth(delta_label, "Helvetica-Bold", 10)
    dpw = dtw + 20
    _pdm_rounded_card(c, W / 2 - dpw / 2, box_y + taux_h - 46, dpw, 17, delta_color, radius=8)
    c.setFillColor(WHITE)
    c.drawCentredString(W / 2, box_y + taux_h - 46 + 4.5, delta_label)

    # ---- Meilleure progression
    positive = [p for p in ranking if p["up"]]
    if positive:
        best = max(positive, key=lambda p: p["evo"])
        ribbon_y = box_y - 20 - 46
        _pdm_rounded_card(c, 36, ribbon_y, W - 72, 46, RED, radius=10)
        c.setFillColor(WHITE)
        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString(W / 2, ribbon_y + 46 / 2 + 4,
                             f"★ MEILLEURE PROGRESSION DE LA SEMAINE : {best['name']} (▲ {best['evo']:.1f}%)")

    # ---- Footer
    footer_h = 34
    c.setFillColor(BLACK)
    c.rect(0, 0, W, footer_h, fill=1, stroke=0)
    c.setFillColor(RED)
    c.rect(0, footer_h - 3, W, 3, fill=1, stroke=0)
    today = datetime.date.today().strftime("%d/%m/%Y")
    c.setFont("Helvetica", 9)
    c.setFillColor(WHITE)
    c.drawCentredString(W / 2, footer_h / 2 - 3,
                         f"Semaine {week}  ·  Généré le {today}  ·  Merci pour votre engagement au quotidien !")

    c.save()
    return len(ranking)
