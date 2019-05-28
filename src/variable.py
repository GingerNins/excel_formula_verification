from dataclasses import dataclass
from openpyxl.cell import Cell
from typing import Union
import re


def parse_value(value: str) -> str:
    if value.startswith('='):
        return value[1:]
    else:
        return value


@dataclass
class Variable(object):
    name: str = None                    # Defined name or coordinate
    sheet: str = None                   # Sheet record is located on
    cell: Union[Cell, None] = None      # Cell record is in, or str for global items
    coordinate: str = None              # Cell coordinate in A1 format
    row: int = None                     # Row value as an integer
    col: int = None                     # Column value as integer
    value: str = None                   # Formula value of the cell/name
    output: Union[str, None] = None     # Evaluated value of the cell/name

    def __post_init__(self):
        if isinstance(self.cell, Cell):
            self.__get_location_information()

        if self.value:
            self.value = parse_value(str(self.value))
        else:
            self.value = parse_value(str(self.cell.value))

    def __get_location_information(self):
        self.coordinate = self.cell.coordinate
        self.row = self.cell.row
        self.col = self.cell.column

    def set_name(self, named_ranges):
        self.name = next((n.name for n in named_ranges if n == self), self.coordinate)

    def set_output(self, wb):
        if self.row:
            self.output = str(wb[self.sheet].cell(row=self.row, column=self.col).value)
        else:
            self.output = None

    def __eq__(self, other):
        """
        Variables are equivalent if they both occupy the same cell in the sheet
        :param other: comparison variable
        :return: True if equivalent
        """
        if self.cell and other.cell:
            return self.cell == other.cell
        return self.name == other.name


@dataclass
class Name(Variable):
    scope: str = None
    is_global: bool = False
    is_used: bool = None

    def __post_init__(self):
        super(Name, self).__post_init__()

    def set_is_used(self, items):
        self.is_used = self.name in [v.name for i in items for v in i.variables]


@dataclass
class Formula(Variable):
    built_ins: list = None
    formats: list = None
    has_digits: bool = None
    in_table: bool = None
    variables: list = None

    def __post_init__(self):
        self.__parse_function()
        super(Formula, self).__post_init__()

    def __parse_function(self):
        self.in_table = True if '[' in self.value else False

        formula_parts = re.split(r'[=*/\-+(),"\s]', self.value)

        # Built in Excel formulas are always in caps
        self.built_ins = [b for b in formula_parts if b.isupper()]

        for part in formula_parts:
            if part.isupper() or part in ['', ' ', '&', "'"]:
                continue
            elif not self.has_digits and part.isdigit():
                self.has_digits = True
            # TODO: If ever have a table formula with TEXT, will need to update this
            elif '#' in part and 'TEXT' in self.value:
                if not self.formats:
                    self.formats = [part]
                else:
                    self.formats.append(part)
            else:
                # Remaining parts should all be names
                # (next((n for n in named_ranges if n['name'] == part), None))
                if not self.variables:
                    self.variables = [part]
                else:
                    self.variables.append(part)

    def update_variables(self, vars):
        self.variables = [v for v in vars if v.name in self.variables]

    def __eq__(self, other):
        return super(Formula, self).__eq__(other)

