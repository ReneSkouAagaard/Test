from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


OUTPUT = Path(__file__).resolve().parents[1] / "pensionsalder_model.xlsx"
MAX_YEARS = 121

DEPOTS = [
    {"name": "Midler i holdingselskab", "row": 41, "balance": 2_000_000, "gain_tax": 0.22, "payout_tax": 0.42, "monthly": 0, "years": 0, "start_offset": 38, "payout_years": 120},
    {"name": "Midler paa pension", "row": 42, "balance": 3_000_000, "gain_tax": 0.153, "payout_tax": 0.38, "monthly": 0, "years": 0, "start_offset": 5, "payout_years": 120},
    {"name": "Frie midler paa aktiedepot", "row": 43, "balance": 1_500_000, "gain_tax": 0.42, "payout_tax": 0.0, "monthly": 0, "years": 0, "start_offset": 38, "payout_years": 120},
    {"name": "Frie midler paa aktiesparekonto", "row": 44, "balance": 500_000, "gain_tax": 0.17, "payout_tax": 0.0, "monthly": 0, "years": 0, "start_offset": 38, "payout_years": 120},
]


def money_style(cell):
    cell.number_format = '#,##0 "kr"'


def pct_style(cell):
    cell.number_format = "0.00%"


def style_header(row):
    for cell in row:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")


def model_ref(cell):
    return f"Model!${cell[0]}${cell[1:]}"


def depot_balance_at_retirement(depot_row, years_cell):
    rate = model_ref(f"J{depot_row}")
    balance = model_ref(f"B{depot_row}")
    annual_payment = model_ref(f"K{depot_row}")
    payment_years = model_ref(f"F{depot_row}")
    n = f"MIN(MAX(0,{years_cell}),{payment_years})"
    return (
        f"({balance}*(1+{rate})^{years_cell}+"
        f"IF({n}<=0,0,IF({rate}=0,{annual_payment}*{n},"
        f"{annual_payment}*((1+{rate})^{n}-1)/{rate})))"
    )


def contribution_formula(depot_row, years_from_now):
    return f"IF({years_from_now}<{model_ref(f'F{depot_row}')},{model_ref(f'K{depot_row}')},0)"


def available_formula(depot_row, future_age):
    start_age = model_ref(f"H{depot_row}")
    payout_years = model_ref(f"I{depot_row}")
    return f"AND({future_age}>={start_age},{future_age}<({start_age}+{payout_years}))"


def main():
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True

    ws = wb.active
    ws.title = "Model"
    proj = wb.create_sheet("Aarlig projektion")
    plan = wb.create_sheet("Udbetalingsplan")

    input_fill = PatternFill("solid", fgColor="FFF2CC")

    ws["A1"] = "Selvpensionsmodel"
    ws["A1"].font = Font(size=18, bold=True)
    ws["A3"] = "Input"
    ws["A3"].font = Font(size=13, bold=True)

    inputs = [
        ("Levealder", 95, "aar"),
        ("Nuvaerende alder", 35, "aar"),
        ("Aarligt forbrug efter skat", 800_000, "kr i dag"),
        ("Lavere aarligt forbrug i sidste aar", 600_000, "kr i dag"),
        ("Antal sidste aar med lavere forbrug", 0, "aar"),
        ("Aarligt afkast REAL foer skat", 0.07, "%"),
        ("Folkepension start", 73, "aar"),
        ("Folkepension foer skat", 7_544, "kr pr maaned i dag"),
        ("Skat paa oevrige indtaegter", 0.38, "%"),
        ("Inflation til nominelle visninger", 0.02, "%"),
        ("Bijob efter selvpension foer skat", 10_000, "kr pr maaned i dag"),
        ("Bijob varighed efter selvpension", 10, "aar"),
        ("Investeringsejendomme foer skat", 0, "kr pr maaned i dag"),
        ("Investeringsejendomme varighed efter selvpension", 120, "aar"),
        ("Referencealder for depotudbetalinger", 73, "aar"),
    ]

    for row, (label, value, unit) in enumerate(inputs, start=4):
        ws[f"A{row}"] = label
        ws[f"B{row}"] = value
        ws[f"C{row}"] = unit
        ws[f"B{row}"].fill = input_fill
        if "%" in unit:
            pct_style(ws[f"B{row}"])
        elif "kr" in unit:
            money_style(ws[f"B{row}"])

    ws["A21"] = "Beregninger"
    ws["A21"].font = Font(size=13, bold=True)
    derived = {
        "A22": "Folkepension foer skat, aarligt",
        "B22": "=B11*12",
        "A23": "Folkepension efter skat, aarligt",
        "B23": "=B22*(1-B12)",
        "A24": "Bijob foer skat, aarligt",
        "B24": "=B14*12",
        "A25": "Bijob efter skat, aarligt",
        "B25": "=B24*(1-B12)",
        "A26": "Investeringsejendomme foer skat, aarligt",
        "B26": "=B16*12",
        "A27": "Investeringsejendomme efter skat, aarligt",
        "B27": "=B26*(1-B12)",
        "A28": "Selvpensionsalder",
        "B28": '=IFERROR(INDEX(\'Aarlig projektion\'!A2:A122,MATCH(TRUE,\'Aarlig projektion\'!I2:I122,0)),"Ikke opnaaet")',
        "A29": "Portefolje ved selvpensionsalder",
        "B29": '=IF(ISNUMBER(B28),INDEX(\'Aarlig projektion\'!C2:C122,MATCH(B28,\'Aarlig projektion\'!A2:A122,0)),"")',
        "A30": "Samlet mangel ved selvpensionsalder",
        "B30": '=IF(ISNUMBER(B28),INDEX(\'Aarlig projektion\'!K2:K122,MATCH(B28,\'Aarlig projektion\'!A2:A122,0)),"")',
        "A31": "Vaegtet realafkast efter skat",
        "B31": "=SUMPRODUCT(B48:B51,J48:J51)/SUM(B48:B51)",
    }
    for cell, value in derived.items():
        ws[cell] = value
    for row in [22, 23, 24, 25, 26, 27, 29, 30]:
        money_style(ws[f"B{row}"])
    pct_style(ws["B31"])

    ws["A32"] = "Oversigt pr. mulig selvpensionsalder"
    ws["A32"].font = Font(size=13, bold=True)
    ws["A33"] = "Alder"
    ws["A34"] = "Portefolje"
    ws["A35"] = "Behov for selvpension"
    ws["A36"] = "Likviditetsmangel"
    for col in range(2, MAX_YEARS + 2):
        col_letter = get_column_letter(col)
        ws[f"{col_letter}33"] = f'=IF($B$5+COLUMN()-2<=$B$4,$B$5+COLUMN()-2,"")'
        ws[f"{col_letter}34"] = f'=IF({col_letter}$33="","",INDEX(\'Aarlig projektion\'!$C:$C,MATCH({col_letter}$33,\'Aarlig projektion\'!$A:$A,0)))'
        ws[f"{col_letter}35"] = f'=IF({col_letter}$33="","",INDEX(\'Aarlig projektion\'!$J:$J,MATCH({col_letter}$33,\'Aarlig projektion\'!$A:$A,0)))'
        ws[f"{col_letter}36"] = f'=IF({col_letter}$33="","",INDEX(\'Aarlig projektion\'!$K:$K,MATCH({col_letter}$33,\'Aarlig projektion\'!$A:$A,0)))'
        money_style(ws[f"{col_letter}34"])
        money_style(ws[f"{col_letter}35"])
        money_style(ws[f"{col_letter}36"])
        ws.column_dimensions[col_letter].width = 14
    style_header(ws[33])
    ws["A46"] = "Depotinput"
    ws["A46"].font = Font(size=13, bold=True)
    depot_headers = [
        "Depot",
        "Startsaldo",
        "Skat paa vaerdistigning",
        "Skat paa udbetaling",
        "Maanedlig indbetaling",
        "Indbetaling varighed aar",
        "Udbetaling starter aar foer referencealder",
        "Udbetaling startalder",
        "Udbetaling varighed aar",
        "Aarligt realafkast efter skat",
        "Aarlig indbetaling",
    ]
    for col, header in enumerate(depot_headers, start=1):
        ws.cell(47, col, header)
    style_header(ws[47])

    for depot in DEPOTS:
        row = depot["row"] + 7
        ws[f"A{row}"] = depot["name"]
        ws[f"B{row}"] = depot["balance"]
        ws[f"C{row}"] = depot["gain_tax"]
        ws[f"D{row}"] = depot["payout_tax"]
        ws[f"E{row}"] = depot["monthly"]
        ws[f"F{row}"] = depot["years"]
        ws[f"G{row}"] = depot["start_offset"]
        ws[f"H{row}"] = f"=$B$18-G{row}"
        ws[f"I{row}"] = depot["payout_years"]
        ws[f"J{row}"] = f"=$B$9*(1-C{row})"
        ws[f"K{row}"] = f"=E{row}*12"
        for cell in [f"B{row}", f"E{row}", f"K{row}"]:
            money_style(ws[cell])
        for cell in [f"C{row}", f"D{row}", f"J{row}"]:
            pct_style(ws[cell])
        for col in range(2, 10):
            ws.cell(row, col).fill = input_fill

    ws["A55"] = "Antagelser"
    ws["A55"].font = Font(size=13, bold=True)
    assumptions = [
        "Alle hovedbeloeb er i realkroner, fordi afkastinput er realt.",
        "De sidste aar med lavere forbrug regnes baglaens fra levealderen inklusiv levealder.",
        "Bijob og investeringsejendomme antages at starte ved den valgte selvpensionsalder og loebe i de valgte antal aar.",
        "Depotudbetalinger kan kun bruges fra den beregnede startalder og inden for den valgte udbetalingsvarighed.",
        "Udbetalinger prioriteres ASK, aktiedepot, holding og pension; resterende saldo investeres fortsat efter depotets afkastskat.",
        "Behov for selvpension i oversigten er et enkelt nutidsvaerdiestimat baseret paa vaegtet realafkast efter skat.",
        "Likviditetsmangel viser, om depoternes udbetalingsregler giver underskud undervejs selv om samlet portefolje er stor nok.",
    ]
    for row, text in enumerate(assumptions, start=56):
        ws[f"A{row}"] = text

    for col, width in {"A": 46, "B": 18, "C": 22, "D": 20, "E": 20, "F": 22, "G": 28, "H": 20, "I": 22, "J": 22, "K": 20}.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A4"

    headers = [
        "Alder",
        "Aar fra nu",
        "Samlet portefolje ved selvpension",
        "Forbrug efter skat i foerste selvpensionsaar",
        "Indkomst efter skat i foerste selvpensionsaar",
        "Nettoforbrug fra depoter i foerste selvpensionsaar",
        "Udbetalingsskat i foerste selvpensionsaar",
        "Slutsaldo ved levealder",
        "Kan gaa paa selvpension",
        "Behov for selvpension",
        "Samlet likviditetsmangel",
        "Folkepension nominelt foer skat",
        "Bijob nominelt foer skat i foerste selvpensionsaar",
        "Investeringsejendomme nominelt foer skat i foerste selvpensionsaar",
    ]
    for col, header in enumerate(headers, start=1):
        proj.cell(1, col, header)
    style_header(proj[1])

    last_row = MAX_YEARS + 1
    helper_start = 15
    helper_width = 14
    helper_need_cols = []
    helper_unmet_cols = []
    helper_end_total_cols = []

    depot_rows_by_name = {depot["name"]: depot["row"] + 7 for depot in DEPOTS}
    withdrawal_rows = [
        depot_rows_by_name["Frie midler paa aktiesparekonto"],
        depot_rows_by_name["Frie midler paa aktiedepot"],
        depot_rows_by_name["Midler i holdingselskab"],
        depot_rows_by_name["Midler paa pension"],
    ]

    for offset in range(MAX_YEARS):
        col = helper_start + offset * helper_width
        letters = [get_column_letter(col + i) for i in range(helper_width)]
        for i, header in enumerate(
            [
                f"Behov +{offset}",
                f"Start holding +{offset}",
                f"Start pension +{offset}",
                f"Start aktiedepot +{offset}",
                f"Start ASK +{offset}",
                f"Gross ASK +{offset}",
                f"Gross aktiedepot +{offset}",
                f"Gross holding +{offset}",
                f"Gross pension +{offset}",
                f"End holding +{offset}",
                f"End pension +{offset}",
                f"End aktiedepot +{offset}",
                f"End ASK +{offset}",
                f"Mangel +{offset}",
            ],
            start=col,
        ):
            proj.cell(1, i, header)
            proj.column_dimensions[get_column_letter(i)].hidden = True
        helper_need_cols.append(letters[0])
        helper_unmet_cols.append(letters[13])
        helper_end_total_cols.append(letters[9:13])

    first_helper_col = get_column_letter(helper_start)
    last_helper_col = get_column_letter(helper_start + MAX_YEARS * helper_width - 1)

    for row in range(2, last_row + 1):
        proj[f"A{row}"] = f'=IF(Model!$B$5+ROW()-2<=Model!$B$4,Model!$B$5+ROW()-2,"")'
        proj[f"B{row}"] = f'=IF(A{row}="","",A{row}-Model!$B$5)'
        starting_balances = [depot_balance_at_retirement(depot["row"] + 7, f"B{row}") for depot in DEPOTS]
        proj[f"C{row}"] = f'=IF(A{row}="","",SUM({",".join(starting_balances)}))'
        proj[f"D{row}"] = f'=IF(A{row}="","",IF(A{row}>Model!$B$4-Model!$B$8,Model!$B$7,Model!$B$6))'
        proj[f"E{row}"] = f'=IF(A{row}="","",IF(A{row}>=Model!$B$10,Model!$B$23,0)+IF(Model!$B$15>0,Model!$B$25,0)+IF(Model!$B$17>0,Model!$B$27,0))'
        proj[f"F{row}"] = f'=IF(A{row}="","",MAX(0,D{row}-E{row}))'

        first_gross_col = get_column_letter(helper_start + 5)
        last_gross_col = get_column_letter(helper_start + 8)
        need_pv_terms = ",".join(
            f"{col}{row}/(1+Model!$B$31)^{offset}"
            for offset, col in enumerate(helper_need_cols)
        )
        unmet_cells = ",".join(f"{col}{row}" for col in helper_unmet_cols)
        life_offset = f"MIN({MAX_YEARS - 1},MAX(0,Model!$B$4-A{row}))"
        helper_range = f"{first_helper_col}{row}:{last_helper_col}{row}"
        proj[f"G{row}"] = f'=IF(A{row}="","",MAX(0,SUM({first_gross_col}{row}:{last_gross_col}{row})-F{row}))'
        proj[f"H{row}"] = (
            f'=IF(A{row}="","",SUM('
            f'INDEX({helper_range},1,{life_offset}*{helper_width}+10),'
            f'INDEX({helper_range},1,{life_offset}*{helper_width}+11),'
            f'INDEX({helper_range},1,{life_offset}*{helper_width}+12),'
            f'INDEX({helper_range},1,{life_offset}*{helper_width}+13)))'
        )
        proj[f"I{row}"] = f'=IF(A{row}="","",SUM({unmet_cells})=0)'
        proj[f"J{row}"] = f'=IF(A{row}="","",SUM({need_pv_terms}))'
        proj[f"K{row}"] = f'=IF(A{row}="","",SUM({unmet_cells}))'
        proj[f"L{row}"] = f'=IF(A{row}="","",IF(A{row}>=Model!$B$10,Model!$B$22*(1+Model!$B$13)^B{row},0))'
        proj[f"M{row}"] = f'=IF(A{row}="","",IF(Model!$B$15>0,Model!$B$24*(1+Model!$B$13)^B{row},0))'
        proj[f"N{row}"] = f'=IF(A{row}="","",IF(Model!$B$17>0,Model!$B$26*(1+Model!$B$13)^B{row},0))'

        for offset in range(MAX_YEARS):
            col = helper_start + offset * helper_width
            need_col, holding_start, pension_start, aktie_start, ask_start = [get_column_letter(col + i) for i in range(5)]
            gross_ask, gross_aktie, gross_holding, gross_pension = [get_column_letter(col + i) for i in range(5, 9)]
            end_holding, end_pension, end_aktie, end_ask, unmet_col = [get_column_letter(col + i) for i in range(9, 14)]
            future_age = f"(A{row}+{offset})"
            years_from_now = f"(B{row}+{offset})"
            target = f"IF({future_age}>Model!$B$4-Model!$B$8,Model!$B$7,Model!$B$6)"
            public_income = f"IF({future_age}>=Model!$B$10,Model!$B$23,0)"
            job_income = f"IF({offset}<Model!$B$15,Model!$B$25,0)"
            property_income = f"IF({offset}<Model!$B$17,Model!$B$27,0)"
            proj[f"{need_col}{row}"] = f'=IF(A{row}="","",IF({future_age}>Model!$B$4,0,MAX(0,{target}-{public_income}-{job_income}-{property_income})))'

            start_cells = {
                depot_rows_by_name["Midler i holdingselskab"]: holding_start,
                depot_rows_by_name["Midler paa pension"]: pension_start,
                depot_rows_by_name["Frie midler paa aktiedepot"]: aktie_start,
                depot_rows_by_name["Frie midler paa aktiesparekonto"]: ask_start,
            }
            end_cells = {
                depot_rows_by_name["Midler i holdingselskab"]: end_holding,
                depot_rows_by_name["Midler paa pension"]: end_pension,
                depot_rows_by_name["Frie midler paa aktiedepot"]: end_aktie,
                depot_rows_by_name["Frie midler paa aktiesparekonto"]: end_ask,
            }
            gross_cells = {
                depot_rows_by_name["Frie midler paa aktiesparekonto"]: gross_ask,
                depot_rows_by_name["Frie midler paa aktiedepot"]: gross_aktie,
                depot_rows_by_name["Midler i holdingselskab"]: gross_holding,
                depot_rows_by_name["Midler paa pension"]: gross_pension,
            }

            for depot in DEPOTS:
                depot_row = depot["row"] + 7
                start_cell = start_cells[depot_row]
                if offset == 0:
                    base = depot_balance_at_retirement(depot_row, f"B{row}")
                else:
                    prev_end = get_column_letter(col - helper_width + list(end_cells).index(depot_row) + 9)
                    base = f"{prev_end}{row}"
                proj[f"{start_cell}{row}"] = f'=IF(A{row}="","",{base}+{contribution_formula(depot_row, years_from_now)})'

            remaining = f"{need_col}{row}"
            for depot_row in withdrawal_rows:
                gross_cell = gross_cells[depot_row]
                start_cell = start_cells[depot_row]
                tax = model_ref(f"D{depot_row}")
                available = available_formula(depot_row, future_age)
                proj[f"{gross_cell}{row}"] = (
                    f'=IF(A{row}="","",IF({available},'
                    f'MIN({start_cell}{row},MAX(0,{remaining})/MAX(0.000001,1-{tax})),0))'
                )
                remaining = f"MAX(0,{remaining}-{gross_cell}{row}*(1-{tax}))"

            for depot in DEPOTS:
                depot_row = depot["row"] + 7
                start_cell = start_cells[depot_row]
                end_cell = end_cells[depot_row]
                gross_cell = gross_cells.get(depot_row)
                if gross_cell is None:
                    raise RuntimeError("Missing gross cell")
                rate = model_ref(f"J{depot_row}")
                proj[f"{end_cell}{row}"] = f'=IF(A{row}="","",MAX(0,{start_cell}{row}-{gross_cell}{row})*(1+{rate}))'
            proj[f"{unmet_col}{row}"] = f'=IF(A{row}="","",{remaining})'

    for col in range(1, 15):
        proj.column_dimensions[get_column_letter(col)].width = 18
    for row in range(2, last_row + 1):
        for col in [3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14]:
            money_style(proj.cell(row, col))
    proj.freeze_panes = "A2"
    proj.auto_filter.ref = f"A1:N{last_row}"

    chart = LineChart()
    chart.title = "Portefolje og slutsaldo"
    chart.y_axis.title = "Kr"
    chart.x_axis.title = "Alder"
    data = Reference(proj, min_col=3, max_col=8, min_row=1, max_row=last_row)
    cats = Reference(proj, min_col=1, min_row=2, max_row=last_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 12
    chart.width = 24
    proj.add_chart(chart, "L2")

    plan["A1"] = "Udbetalingsplan ved selvpensionsalder"
    plan["A1"].font = Font(size=16, bold=True)
    plan["A3"] = "Selvpensionsalder"
    plan["B3"] = "=Model!B28"
    plan["A4"] = "Startportefolje i plan"
    plan["B4"] = '=IF(ISNUMBER(B3),INDEX(\'Aarlig projektion\'!C:C,MATCH(B3,\'Aarlig projektion\'!A:A,0)),"")'
    money_style(plan["B4"])
    plan["A5"] = "Bemerkning"
    plan["B5"] = "Planen viser den simulerede brug af hvert depot ved selvpensionsalderen."

    plan_headers = [
        "Alder",
        "Holding start",
        "Pensionsdepot start",
        "Aktiedepot start",
        "ASK start",
        "Forbrug efter skat",
        "Indkomst efter skat",
        "Netto fra depoter",
        "Brutto ASK",
        "Brutto aktiedepot",
        "Brutto holding",
        "Brutto pension",
        "Mangel",
        "Slutsaldo samlet",
    ]
    for col, header in enumerate(plan_headers, start=1):
        plan.cell(7, col, header)
    style_header(plan[7])

    for row in range(8, last_row + 7):
        offset = row - 8
        if row == 8:
            plan[f"A{row}"] = '=IF(ISNUMBER($B$3),$B$3,"")'
        else:
            prev = row - 1
            plan[f"A{row}"] = f'=IF(OR(A{prev}="",A{prev}>=Model!$B$4),"",A{prev}+1)'

        source_row = 'MATCH($B$3,\'Aarlig projektion\'!$A:$A,0)'
        source_col = helper_start + offset * helper_width
        for plan_col, source_offset in zip(range(2, 14), [1, 2, 3, 4, None, None, None, 5, 6, 7, 8, 13]):
            col_letter = get_column_letter(plan_col)
            if source_offset is None:
                continue
            source_letter = get_column_letter(source_col + source_offset)
            plan[f"{col_letter}{row}"] = f'=IF(A{row}="","",INDEX(\'Aarlig projektion\'!{source_letter}:{source_letter},{source_row}))'

        future_age = f"A{row}"
        target = f"IF({future_age}>Model!$B$4-Model!$B$8,Model!$B$7,Model!$B$6)"
        public_income = f"IF({future_age}>=Model!$B$10,Model!$B$23,0)"
        job_income = f"IF({offset}<Model!$B$15,Model!$B$25,0)"
        property_income = f"IF({offset}<Model!$B$17,Model!$B$27,0)"
        plan[f"F{row}"] = f'=IF(A{row}="","",{target})'
        plan[f"G{row}"] = f'=IF(A{row}="","",{public_income}+{job_income}+{property_income})'
        plan[f"H{row}"] = f'=IF(A{row}="","",MAX(0,F{row}-G{row}))'
        end_total_parts = [
            f"INDEX('Aarlig projektion'!{get_column_letter(source_col + i)}:{get_column_letter(source_col + i)},{source_row})"
            for i in range(9, 13)
        ]
        plan[f"N{row}"] = f'=IF(A{row}="","",SUM({",".join(end_total_parts)}))'

    for col in range(1, 15):
        plan.column_dimensions[get_column_letter(col)].width = 18
    for row in range(8, last_row + 7):
        for col in range(2, 15):
            money_style(plan.cell(row, col))
    plan.freeze_panes = "A8"
    plan.auto_filter.ref = f"A7:N{last_row + 6}"

    wb.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
