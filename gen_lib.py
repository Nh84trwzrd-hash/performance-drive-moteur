import openpyxl, re, datetime
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart.data_source import StrRef, AxDataSource


def get_week_and_employees(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Preparation']
    periode = ws.cell(row=5, column=3).value  # "Du DD/MM/YYYY Au DD/MM/YYYY"
    m = re.search(r'Du (\d{2})/(\d{2})/(\d{4})', periode)
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    dt = datetime.date(y, mo, d)
    iso = dt.isocalendar()
    week = iso[1]

    row = 25
    employees = []  # (name, prep_row, articles_row)
    while row <= ws.max_row:
        name = ws.cell(row=row, column=2).value
        if isinstance(name, str) and '(' in name and ')' in name:
            employees.append((name.strip(), row+1, row+3))
            row += 12
        else:
            row += 1
    return week, employees


def build(path, outpath, taux_actuelle=None, taux_precedente=None):
    week, employees = get_week_and_employees(path)
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
    blue_font = Font(name=arial, color='0000FF')

    for i, (name, prep_row, art_row) in enumerate(employees):
        r = first_data_row + i
        ws.cell(row=r, column=1, value=name).font = Font(name=arial)
        cb = ws.cell(row=r, column=2, value=None)
        cb.fill = yellow_fill
        cb.font = blue_font
        cc = ws.cell(row=r, column=3, value=f"='{prep_sheet_name}'!C{art_row}")
        cc.font = Font(name=arial, color='008000')
        cd = ws.cell(row=r, column=4, value=f"='{prep_sheet_name}'!C{prep_row}")
        cd.font = Font(name=arial, color='008000')
        ce = ws.cell(row=r, column=5, value=f"=IFERROR(C{r}/B{r},\"\")")
        ce.font = Font(name=arial)
        ce.number_format = '0.00'
        cf = ws.cell(row=r, column=6, value=f"=IFERROR(E{r}/60,\"\")")
        cf.font = Font(name=arial)
        cf.number_format = '0.00'
        cg = ws.cell(row=r, column=7, value=f"=IFERROR(D{r}/B{r},\"\")")
        cg.font = Font(name=arial)
        cg.number_format = '0.00'
        ch = ws.cell(row=r, column=8, value=None)
        ch.fill = yellow_fill
        ch.font = blue_font
        ch.number_format = '0.00'
        ci = ws.cell(row=r, column=9,
                      value=f"=IFERROR(IF(OR(E{r}=\"\",H{r}=\"\"),\"\",IF(E{r}>=H{r},\"▲ \",\"▼ \")&ROUND(ABS(E{r}-H{r})/H{r}*100,1)&\"%\"),\"\")")
        ci.font = Font(name=arial, bold=True)
        ci.alignment = Alignment(horizontal='center')

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

    wb.save(outpath)
    return week, n
