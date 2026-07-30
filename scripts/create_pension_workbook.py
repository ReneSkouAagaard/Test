from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


OUTPUT = Path(__file__).resolve().parents[1] / "pensionsalder_model.xlsx"
MAX_ROWS = 121


def money_style(cell):
    cell.number_format = '#,##0 "kr"'


def pct_style(cell):
    cell.number_format = "0.00%"


def style_header(row):
    for cell in row:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")


def main():
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True

    ws = wb.active
    ws.title = "Model"
    proj = wb.create_sheet("Aarlig projektion")
    plan = wb.create_sheet("Udbetalingsplan")

    # Model sheet
    ws["A1"] = "Pensionsalder model"
    ws["A1"].font = Font(size=18, bold=True)
    ws["A3"] = "Input"
    ws["A3"].font = Font(size=13, bold=True)

    inputs = [
        ("Levealder", 95, "aar"),
        ("Nuvaerende portefolje", 7_000_000, "kr"),
        ("Aarligt forbrug efter skat", 800_000, "kr i dag"),
        ("Aarligt afkast REAL foer skat", 0.07, "%"),
        ("Skat paa afkast", 0.22, "%"),
        ("Skat paa udbetalinger", 0.38, "%"),
        ("Nuvaerende alder", 35, "aar"),
        ("Offentlig pensionsudbetaling start", 73, "aar"),
        ("Sats for offentlig pension foer skat", 7_544, "kr pr maaned i dag"),
        ("Inflation til nominelle visninger", 0.02, "%"),
        ("Bijob efter pensionsalder foer skat", 10_000, "kr pr maaned i dag"),
        ("Bijob varighed efter pensionsalder", 10, "aar"),
    ]

    start_row = 4
    for idx, (label, value, unit) in enumerate(inputs, start=start_row):
        ws[f"A{idx}"] = label
        ws[f"B{idx}"] = value
        ws[f"C{idx}"] = unit
        if "%" in unit:
            pct_style(ws[f"B{idx}"])
        elif "kr" in unit:
            money_style(ws[f"B{idx}"])

    ws["A17"] = "Beregninger"
    ws["A17"].font = Font(size=13, bold=True)
    derived = {
        "A18": "Aarligt realafkast efter skat",
        "B18": "=B7*(1-B8)",
        "A19": "Offentlig pension foer skat, aarligt",
        "B19": "=B12*12",
        "A20": "Offentlig pension efter skat, aarligt",
        "B20": "=B19*(1-B9)",
        "A21": "Bijob foer skat, aarligt",
        "B21": "=B14*12",
        "A22": "Bijob efter skat, aarligt",
        "B22": "=B21*(1-B9)",
        "A23": "Bijob varighed",
        "B23": "=B15",
        "A24": "Optimal pensionsalder",
        "B24": '=IFERROR(INDEX(\'Aarlig projektion\'!A2:A122,MATCH(TRUE,\'Aarlig projektion\'!K2:K122,0)),"Ikke opnaaet")',
        "A25": "Portefolje ved optimal alder",
        "B25": '=IF(ISNUMBER(B24),INDEX(\'Aarlig projektion\'!C2:C122,MATCH(B24,\'Aarlig projektion\'!A2:A122,0)),"")',
        "A26": "Behov ved optimal alder",
        "B26": '=IF(ISNUMBER(B24),INDEX(\'Aarlig projektion\'!D2:D122,MATCH(B24,\'Aarlig projektion\'!A2:A122,0)),"")',
        "A27": "Margin ved optimal alder",
        "B27": '=IF(ISNUMBER(B24),B25-B26,"")',
    }
    for cell, value in derived.items():
        ws[cell] = value
    pct_style(ws["B18"])
    for row in range(19, 28):
        if row not in [23, 24]:
            money_style(ws[f"B{row}"])

    ws["A29"] = "Antagelser"
    ws["A29"].font = Font(size=13, bold=True)
    assumptions = [
        "Alle hovedbeloeb er i realkroner, fordi afkastinput er realt.",
        "Offentlig pension antages fuldt inflationsjusteret, saa dens reale koebekraft er konstant.",
        "Bijob antages at starte ved pensionsalderen, loebe i det valgte antal aar og have konstant real koebekraft.",
        "Portefoljebehovet er brutto foer udbetalingsskat og beregnes, saa portefoljen er 0 kr ved slutningen af levealderen.",
        "Udbetalinger antages at ske ved starten af hvert pensionsaar; resterende saldo investeres fortsat.",
    ]
    for i, text in enumerate(assumptions, start=30):
        ws[f"A{i}"] = text

    for col, width in {"A": 44, "B": 18, "C": 28}.items():
        ws.column_dimensions[col].width = width

    input_fill = PatternFill("solid", fgColor="FFF2CC")
    for row in range(4, 16):
        ws[f"B{row}"].fill = input_fill
    ws.freeze_panes = "A4"

    # Projection sheet
    headers = [
        "Alder",
        "Aar fra nu",
        "Forventet portefolje",
        "Nodvendig portefolje",
        "Offentlig pension efter skat",
        "Bijob efter skat i foerste pensionsaar",
        "Samlet indkomst efter skat i foerste pensionsaar",
        "Nettoforbrug fra portefolje",
        "Brutto udbetaling fra portefolje",
        "Margin",
        "Kan pensioneres",
        "Offentlig pension nominelt foer skat",
        "Bijob nominelt foer skat i foerste pensionsaar",
    ]
    for col, header in enumerate(headers, start=1):
        proj.cell(1, col, header)
    style_header(proj[1])

    last_row = MAX_ROWS + 1
    for row in range(2, last_row + 1):
        age_formula = f'=IF(Model!$B$10+ROW()-2<=Model!$B$4,Model!$B$10+ROW()-2,"")'
        proj[f"A{row}"] = age_formula
        proj[f"B{row}"] = f'=IF(A{row}="","",A{row}-Model!$B$10)'
        proj[f"C{row}"] = f'=IF(A{row}="","",Model!$B$5*(1+Model!$B$18)^B{row})'
        proj[f"E{row}"] = f'=IF(A{row}="","",IF(A{row}>=Model!$B$11,Model!$B$20,0))'
        proj[f"F{row}"] = f'=IF(A{row}="","",IF(Model!$B$15>0,Model!$B$22,0))'
        proj[f"G{row}"] = f'=IF(A{row}="","",E{row}+F{row})'
        proj[f"H{row}"] = f'=IF(A{row}="","",MAX(0,Model!$B$6-G{row}))'
        proj[f"I{row}"] = f'=IF(A{row}="","",H{row}/(1-Model!$B$9))'
        future_ages = f"$A$2:$A${last_row}"
        public_income = f"IF({future_ages}>=Model!$B$11,Model!$B$20,0)"
        bijob_income = (
            f"IF(({future_ages}>=A{row})*({future_ages}<A{row}+Model!$B$15),"
            f"Model!$B$22,0)"
        )
        net_need = f"(Model!$B$6-{public_income}-{bijob_income})"
        proj[f"D{row}"] = (
            f'=IF(A{row}="","",SUMPRODUCT(({future_ages}>=A{row})'
            f'*({future_ages}<>"")*({net_need}>0)*{net_need}/(1-Model!$B$9)/'
            f'(1+Model!$B$18)^({future_ages}-A{row})))'
        )
        proj[f"J{row}"] = f'=IF(A{row}="","",C{row}-D{row})'
        proj[f"K{row}"] = f'=IF(A{row}="","",C{row}>=D{row})'
        proj[f"L{row}"] = f'=IF(A{row}="","",IF(A{row}>=Model!$B$11,Model!$B$19*(1+Model!$B$13)^B{row},0))'
        proj[f"M{row}"] = f'=IF(A{row}="","",IF(Model!$B$15>0,Model!$B$21*(1+Model!$B$13)^B{row},0))'

    for col in range(1, 14):
        proj.column_dimensions[get_column_letter(col)].width = 18
    for row in range(2, last_row + 1):
        for col in [3, 4, 5, 6, 7, 8, 9, 10, 12, 13]:
            money_style(proj.cell(row, col))
    proj.freeze_panes = "A2"
    proj.auto_filter.ref = f"A1:M{last_row}"

    chart = LineChart()
    chart.title = "Portefolje vs. behov"
    chart.y_axis.title = "Kr"
    chart.x_axis.title = "Alder"
    data = Reference(proj, min_col=3, max_col=4, min_row=1, max_row=last_row)
    cats = Reference(proj, min_col=1, min_row=2, max_row=last_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 12
    chart.width = 24
    proj.add_chart(chart, "L2")

    # Withdrawal plan sheet
    plan["A1"] = "Udbetalingsplan ved optimal pensionsalder"
    plan["A1"].font = Font(size=16, bold=True)
    plan["A3"] = "Valgt pensionsalder"
    plan["B3"] = "=Model!B24"
    plan["A4"] = "Startportefolje i plan"
    plan["B4"] = '=IF(ISNUMBER(B3),INDEX(\'Aarlig projektion\'!D:D,MATCH(B3,\'Aarlig projektion\'!A:A,0)),"")'
    money_style(plan["B4"])
    plan["A5"] = "Bemerkning"
    plan["B5"] = "Planen starter med den nodvendige portefolje, saa saldoen rammer 0 kr ved levealderen."

    plan_headers = [
        "Alder",
        "Startsaldo",
        "Offentlig pension efter skat",
        "Bijob efter skat",
        "Samlet indkomst efter skat",
        "Nettoforbrug fra portefolje",
        "Brutto udbetaling",
        "Saldo efter udbetaling",
        "Afkast efter skat",
        "Slutsaldo",
    ]
    for col, header in enumerate(plan_headers, start=1):
        plan.cell(7, col, header)
    style_header(plan[7])

    for row in range(8, last_row + 7):
        if row == 8:
            plan[f"A{row}"] = '=IF(ISNUMBER($B$3),$B$3,"")'
            plan[f"B{row}"] = '=IF(A8="","",$B$4)'
        else:
            prev = row - 1
            plan[f"A{row}"] = f'=IF(OR(A{prev}="",A{prev}>=Model!$B$4),"",A{prev}+1)'
            plan[f"B{row}"] = f'=IF(A{row}="","",J{prev})'
        plan[f"C{row}"] = f'=IF(A{row}="","",IF(A{row}>=Model!$B$11,Model!$B$20,0))'
        plan[f"D{row}"] = f'=IF(A{row}="","",IF(A{row}<$B$3+Model!$B$15,Model!$B$22,0))'
        plan[f"E{row}"] = f'=IF(A{row}="","",C{row}+D{row})'
        plan[f"F{row}"] = f'=IF(A{row}="","",MAX(0,Model!$B$6-E{row}))'
        plan[f"G{row}"] = f'=IF(A{row}="","",F{row}/(1-Model!$B$9))'
        plan[f"H{row}"] = f'=IF(A{row}="","",MAX(0,B{row}-G{row}))'
        plan[f"I{row}"] = f'=IF(A{row}="","",H{row}*Model!$B$18)'
        plan[f"J{row}"] = f'=IF(A{row}="","",H{row}+I{row})'

    for col in range(1, 11):
        plan.column_dimensions[get_column_letter(col)].width = 18
    plan.column_dimensions["B"].width = 22
    plan.column_dimensions["J"].width = 22
    for row in range(8, last_row + 7):
        for col in range(2, 11):
            money_style(plan.cell(row, col))
    plan.freeze_panes = "A8"
    plan.auto_filter.ref = f"A7:J{last_row + 6}"

    wb.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
