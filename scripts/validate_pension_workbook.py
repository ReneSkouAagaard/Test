import subprocess
import time
from pathlib import Path

import openpyxl
import uno
from com.sun.star.beans import PropertyValue


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


def calculate_with_libreoffice():
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
            uno.systemPathToFileUrl(str(WORKBOOK)),
            "_blank",
            0,
            (prop("Hidden", True), prop("ReadOnly", False)),
        )
        if doc is None:
            raise RuntimeError(f"Could not open {WORKBOOK}")
        doc.calculateAll()
        doc.storeAsURL(
            uno.systemPathToFileUrl(str(CALCULATED)),
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


def main():
    calculate_with_libreoffice()
    wb = openpyxl.load_workbook(CALCULATED, data_only=True)
    projection = wb["Aarlig projektion"]
    model = wb["Model"]

    for cell in ["C2", "D2", "F2", "G2", "H2", "J2", "N2", "AA2", "BMQ2"]:
        assert_numeric(projection, cell)
    optimal_age = assert_numeric(model, "B28")
    if optimal_age < model["B5"].value or optimal_age > model["B4"].value:
        raise AssertionError(f"Optimal age is outside the modeled age range: {optimal_age!r}")
    assert_numeric(model, "B29")
    assert_numeric(model, "B30")

    print("Workbook calculation OK")


if __name__ == "__main__":
    main()
