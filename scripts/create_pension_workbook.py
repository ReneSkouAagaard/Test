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

    ws["A16"] = "Beregninger"
    ws["A16"].font = Font(size=13, bold=True)
    derived = {
        "A17": "Aarligt realafkast efter skat",
        "B17": "=B7*(1-B8)",
        "A18": "Offentlig pension foer skat, aarligt",
        "B18": "=B12*12",
        "A19": "Offentlig pension efter skat, aarligt",
        "B19": "=B18*(1-B9)",
        "A20": "Optimal pensionsalder",
        "B20": '=IFERROR(INDEX(\'Aarlig projektion\'!A2:A122,MATCH(TRUE,\'Aarlig projektion\'!I2:I122,0)),"Ikke opnaaet")',
        "A21": "Portefolje ved optimal alder",
        "B21": '=IF(ISNUMBER(B20),INDEX(\'Aarlig projektion\'!C2:C122,MATCH(B20,\'Aarlig projektion\'!A2:A122,0)),"")',
        "A22": "Behov ved optimal alder",
        "B22": '=IF(ISNUMBER(B20),INDEX(\'Aarlig projektion\'!D2:D122,MATCH(B20,\'Aarlig projektion\'!A2:A122,0)),"")',
        "A23": "Margin ved optimal alder",
        "B23": '=IF(ISNUMBER(B20),B21-B22,"")',
    }
    for cell, value in derived.items():
        ws[cell] = value
    pct_style(ws["B17"])
    for row in range(18, 24):
        if row != 20:
            money_style(ws[f"B{row}"])

    ws["A26"] = "Antagelser"
    ws["A26"].font = Font(size=13, bold=True)
    assumptions = [
        "Alle hovedbeloeb er i realkroner, fordi afkastinput er realt.",
        "Offentlig pension antages fuldt inflationsjusteret, saa dens reale koebekraft er konstant.",
        "Portefoljebehovet er brutto foer udbetalingsskat og beregnes, saa portefoljen er 0 kr ved slutningen af levealderen.",
        "Udbetalinger antages at ske ved starten af hvert pensionsaar; resterende saldo investeres fortsat.",
    ]
    for i, text in enumerate(assumptions, start=27):
        ws[f"A{i}"] = text

    for col, width in {"A": 42, "B": 18, "C": 28}.items():
        ws.column_dimensions[col].width = width

    input_fill = PatternFill("solid", fgColor="FFF2CC")
    for row in range(4, 14):
        ws[f"B{row}"].fill = input_fill
    ws.freeze_panes = "A4"

    # Projection sheet
    headers = [
        "Alder",
        "Aar fra nu",
        "Forventet portefolje",
        "Nodvendig portefolje",
        "Offentlig pension efter skat",
        "Nettoforbrug fra portefolje",
        "Brutto udbetaling fra portefolje",
        "Margin",
        "Kan pensioneres",
        "Offentlig pension nominelt foer skat",
    ]
    for col, header in enumerate(headers, start=1):
        proj.cell(1, col, header)
    style_header(proj[1])

    last_row = MAX_ROWS + 1
    for row in range(2, last_row + 1):
        age_formula = f'=IF(Model!$B$10+ROW()-2<=Model!$B$4,Model!$B$10+ROW()-2,"")'
        proj[f"A{row}"] = age_formula
        proj[f"B{row}"] = f'=IF(A{row}="","",A{row}-Model!$B$10)'
        proj[f"C{row}"] = f'=IF(A{row}="","",Model!$B$5*(1+Model!$B$17)^B{row})'
        proj[f"E{row}"] = f'=IF(A{row}="","",IF(A{row}>=Model!$B$11,Model!$B$19,0))'
        proj[f"F{row}"] = f'=IF(A{row}="","",MAX(0,Model!$B$6-E{row}))'
        proj[f"G{row}"] = f'=IF(A{row}="","",F{row}/(1-Model!$B$9))'
        proj[f"D{row}"] = (
            f'=IF(A{row}="","",SUMPRODUCT(($A$2:$A${last_row}>=A{row})'
            f'*($A$2:$A${last_row}<>"")*$G$2:$G${last_row}/'
            f'(1+Model!$B$17)^($A$2:$A${last_row}-A{row})))'
        )
        proj[f"H{row}"] = f'=IF(A{row}="","",C{row}-D{row})'
        proj[f"I{row}"] = f'=IF(A{row}="","",C{row}>=D{row})'
        proj[f"J{row}"] = f'=IF(A{row}="","",IF(A{row}>=Model!$B$11,Model!$B$18*(1+Model!$B$13)^B{row},0))'

    for col in range(1, 11):
        proj.column_dimensions[get_column_letter(col)].width = 18
    for row in range(2, last_row + 1):
        for col in [3, 4, 5, 6, 7, 8, 10]:
            money_style(proj.cell(row, col))
    proj.freeze_panes = "A2"
    proj.auto_filter.ref = f"A1:J{last_row}"

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
    plan["B3"] = "=Model!B20"
    plan["A4"] = "Startportefolje i plan"
    plan["B4"] = '=IF(ISNUMBER(B3),INDEX(\'Aarlig projektion\'!D:D,MATCH(B3,\'Aarlig projektion\'!A:A,0)),"")'
    money_style(plan["B4"])
    plan["A5"] = "Bemerkning"
    plan["B5"] = "Planen starter med den nodvendige portefolje, saa saldoen rammer 0 kr ved levealderen."

    plan_headers = [
        "Alder",
        "Startsaldo",
        "Offentlig pension efter skat",
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
            plan[f"B{row}"] = f'=IF(A{row}="","",H{prev})'
        plan[f"C{row}"] = f'=IF(A{row}="","",IF(A{row}>=Model!$B$11,Model!$B$19,0))'
        plan[f"D{row}"] = f'=IF(A{row}="","",MAX(0,Model!$B$6-C{row}))'
        plan[f"E{row}"] = f'=IF(A{row}="","",D{row}/(1-Model!$B$9))'
        plan[f"F{row}"] = f'=IF(A{row}="","",MAX(0,B{row}-E{row}))'
        plan[f"G{row}"] = f'=IF(A{row}="","",F{row}*Model!$B$17)'
        plan[f"H{row}"] = f'=IF(A{row}="","",F{row}+G{row})'

    for col in range(1, 9):
        plan.column_dimensions[get_column_letter(col)].width = 18
    plan.column_dimensions["B"].width = 22
    plan.column_dimensions["H"].width = 22
    for row in range(8, last_row + 7):
        for col in range(2, 9):
            money_style(plan.cell(row, col))
    plan.freeze_panes = "A8"
    plan.auto_filter.ref = f"A7:H{last_row + 6}"

    wb.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
