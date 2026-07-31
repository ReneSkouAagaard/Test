import subprocess
import time
from shutil import copyfile
from pathlib import Path

import openpyxl
import uno
from com.sun.star.beans import PropertyValue
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "pensionsalder_model.xlsx"
CALCULATED = Path("/tmp/pension-workbook-calculated.xlsx")
PORT = 2003


def prop(name, value):
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def connect_to_calc():
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    url = f"uno:socket,host=127.0.0.1,port={PORT};urp;StarOffice.ComponentContext"
    last_error = None
    for _ in range(30):
        try:
            return resolver.resolve(url)
        except Exception as exc:
            last_error = exc
            time.sleep(0.5)
    raise RuntimeError(f"Could not connect to LibreOffice: {last_error}")


def calculate_with_libreoffice(workbook=WORKBOOK, calculated=CALCULATED):
    cmd = [
        "soffice",
        "--headless",
        f"--accept=socket,host=127.0.0.1,port={PORT};urp;StarOffice.ComponentContext",
        "--norestore",
        "--nofirststartwizard",
    ]
    proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        ctx = connect_to_calc()
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx
        )
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(workbook)),
            "_blank",
            0,
            (prop("Hidden", True), prop("ReadOnly", False)),
        )
        if doc is None:
            raise RuntimeError(f"Could not open {workbook}")
        doc.calculateAll()
        doc.storeAsURL(
            uno.systemPathToFileUrl(str(calculated)),
            (prop("FilterName", "Calc MS Excel 2007 XML"), prop("Overwrite", True)),
        )
        doc.close(True)
        desktop.terminate()
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()


def assert_numeric(ws, cell):
    value = ws[cell].value
    if not isinstance(value, (int, float)):
        raise AssertionError(f"{ws.title}!{cell} is not numeric after calculation: {value!r}")
    return value


def calculated_j2_with_change(cell, value):
    variant = Path(f"/tmp/pension-workbook-{cell.lower()}-{value}.xlsx")
    calculated = Path(f"/tmp/pension-workbook-{cell.lower()}-{value}-calculated.xlsx")
    copyfile(WORKBOOK, variant)
    wb = openpyxl.load_workbook(variant)
    wb["Model"][cell] = value
    wb.save(variant)
    calculate_with_libreoffice(variant, calculated)
    result = openpyxl.load_workbook(calculated, data_only=True)
    return assert_numeric(result["Aarlig projektion"], "J2")


def main():
    calculate_with_libreoffice()
    wb = openpyxl.load_workbook(CALCULATED, data_only=True)
    projection = wb["Aarlig projektion"]
    model = wb["Model"]

    helper_start = 15
    helper_width = 15
    last_unmet = f"{get_column_letter(helper_start + 120 * helper_width + 14)}2"
    for cell in ["C2", "D2", "F2", "G2", "H2", "J2", "K2", "O2", "AC2", last_unmet]:
        assert_numeric(projection, cell)
    for cell in ["B34", "B35", "B36"]:
        assert_numeric(model, cell)
    optimal_age = assert_numeric(model, "B28")
    if optimal_age < model["B5"].value or optimal_age > model["B4"].value:
        raise AssertionError(f"Optimal age is outside the modeled age range: {optimal_age!r}")
    assert_numeric(model, "B29")
    assert_numeric(model, "B30")

    life_offset = int(model["B4"].value - projection["A2"].value)
    end_cols = [helper_start + life_offset * helper_width + i for i in range(10, 14)]
    expected_h2 = sum(projection[f"{get_column_letter(col)}2"].value for col in end_cols)
    actual_h2 = assert_numeric(projection, "H2")
    if abs(actual_h2 - expected_h2) > 1:
        raise AssertionError(f"Aarlig projektion!H2 should use balances at life expectancy, got {actual_h2}, expected {expected_h2}")

    baseline_j2 = assert_numeric(projection, "J2")
    if calculated_j2_with_change("B8", 5) == baseline_j2:
        raise AssertionError("Changing Model!B8 did not change Aarlig projektion!J2")
    if calculated_j2_with_change("B16", 10_000) == baseline_j2:
        raise AssertionError("Changing Model!B16 did not change Aarlig projektion!J2")

    print("Workbook calculation OK")


if __name__ == "__main__":
    main()
