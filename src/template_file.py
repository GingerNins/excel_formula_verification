import openpyxl
from openpyxl import Workbook, utils
import pandas as pd
from variable import Formula, Name, Variable


def process_template_file(filename: str, output_file: str = None) -> dict:
    """
    Processes the Excel file aggregating all formulas, named ranges, constants
    matching them together, storing the list of formulas and list of names in a dict
    :param filename: name of template file
    :param output_file: an excel file of all formulas and named ranges (optional)
    :return: dict of formula list and named ranges list
    """
    # Parses out formulas and names
    try:
        wb: Workbook = openpyxl.load_workbook(filename)
    except PermissionError as e:
        print(e)
        exit(1)

    named_ranges: list = _get_named_ranges(wb)
    formulas, constants = _get_formulas_and_constants(wb)

    # Parses out the actual values instead of the formulas
    # and pairs them with the original formula/named range record
    try:
        wb_data: Workbook = openpyxl.load_workbook(filename, data_only=True)
    except PermissionError as e:
        print(e)
        exit(1)

    for c in constants:
        c.set_name(named_ranges)
        c.set_output(wb_data)

    for f in formulas:
        f.set_name(named_ranges)
        f.set_output(wb_data)
        f.update_variables(named_ranges)

    for n in named_ranges:
        n.set_is_used(formulas)
        n.set_output(wb_data)

    wb.close()
    wb_data.close()

    if output_file:
        for name, items in zip(['named_ranges', 'formulas', 'constants'],
                               [named_ranges, formulas, constants]):
            output_formulas_to_excel(output_file, name, items)

    return {'formulas': formulas, 'names': named_ranges, 'constants': constants}


def _get_named_ranges(wb) -> list:
    """
    Aggregates all named ranges into a list
    :return: list of all named ranges
    """
    named_range_list = []

    for dn in wb.defined_names.definedName:
        name = dn.name
        scope = wb.sheetnames[dn.localSheetId] if dn.localSheetId is not None else 'Workbook'
        dest = wb.defined_names[name].destinations

        for sheet_name, rng in dest:
            ws = wb[sheet_name]

            # If the range is only a single cell
            if not isinstance(ws[rng], tuple):
                named_range_list.append(
                    Name(sheet=sheet_name, name=name, scope=scope, cell=ws[rng]))
                break
            # If the named range contains multiple cells
            else:
                for c in ws[rng]:
                    named_range_list.append(
                        Name(sheet=sheet_name, name=name, scope=scope, cell=c[0]))
                break
        else:
            # Global Constants
            named_range_list.append(
                Name(name=name, scope=scope, value=dn.attr_text, is_global=True))

    return named_range_list


def _get_formulas_and_constants(wb) -> tuple:
    """
    Aggregates all formulas and variables in the workbook in a list
    :return: a list of Formulas and Variables
    """
    formula_list = []
    constants_list = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        last_cell = utils.get_column_letter(ws.max_column) + str(ws.max_row)
        working_range = ws['A1':last_cell]

        for row in working_range:
            for cell in row:
                # Skip the formula if:
                # - It's blank
                # - It's a string and doesn't start with "=" (indicates just text)
                if cell.value is None or (isinstance(cell.value, str) and not cell.value.startswith('=')):
                    continue

                if isinstance(cell.value, str) and cell.value.startswith('='):
                    value = cell.value[1:]
                    formula_list.append(Formula(sheet=sheet_name, cell=cell, value=value))
                else:
                    value = cell.value
                    constants_list.append(Variable(sheet=sheet_name, cell=cell, value=value))

    return formula_list, constants_list


def output_formulas_to_excel(outfile, name, items):
    try:
        book = openpyxl.load_workbook(outfile)
    except FileNotFoundError:
        book = openpyxl.Workbook()
        sheet = book.get_sheet_by_name('Sheet')
        book.remove_sheet(sheet)

    writer = pd.ExcelWriter(outfile, engine='openpyxl')
    writer.book = book

    df = pd.DataFrame([i.__dict__ for i in items])
    df.to_excel(writer, sheet_name=name)
    writer.save()
