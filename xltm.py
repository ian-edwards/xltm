from pathlib import Path
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree
from xml.etree.ElementTree import Element
from functools import reduce
import sys
import csv
import re

# Images
def write_images(path: str, zip: ZipFile):
    for (id, data, ext) in book_images(zip):
        Path(path).joinpath(Path(f'{id}{ext}')).write_bytes(data)

def is_image(element: Element):
    return element.get('Type').endswith('/image')

def image_id(element: Element):
    return int(element.get('Id')[3:]) - 1

def image_path(element: Element):
    return 'xl' + element.get('Target')[2:]

def image_elements(zip: ZipFile):
    return [e for e in xml_elements('Relationship', read_file('xl/richData/_rels/richValueRel.xml.rels', zip)) if is_image(e)]

def images_info(zip: ZipFile):
    return [(image_path(e), image_id(e)) for e in image_elements(zip)]

def book_images(zip: ZipFile):
    return [(i, zip.read(p), Path(p).suffix) for p, i in images_info(zip)]

# Sheets
def write_sheets(path: str, zip: ZipFile):
    for (name, cells) in book_sheets(zip).items():
        with Path(path).joinpath(Path(f'{name}.csv')).open('w', newline='') as file:
            csv.writer(file).writerows(cells)

def book_sheets(zip: ZipFile):
    return {sheet_name(e) : cells_csv(sheet_cells(sheet_path(e), zip)) for e in xml_elements('sheet', read_file('xl/workbook.xml', zip))}

def cells_csv(cells: dict[(int, int), int]):
    return [[cells.get((r, c)) for c in range(columns_count(cells))] for r in range(rows_count(cells))]

def sheet_cells(path: str, zip: ZipFile):
    return {cell_coordinates(c) : cell_image(c) for c in xml_elements('c', read_file(path, zip))}

def sheet_name(sheet: Element):
    return sheet.get('name')

def sheet_path(sheet: Element):
    return f'xl/worksheets/sheet{sheet.get('sheetId')}.xml'

def cell_coordinates(cell: Element):
    return (cell_row(a := cell.get('r')), cell_column(a))

def cell_image(cell: Element):
    return int(cell.get('vm')) - 1

def cell_row(address: str):
    return int(re.match('[A-Z]+(?P<row>\\d+)', address).group('row')) - 1

def column_id(column: str):
    return reduce(lambda _, n: 26 * n[0] + ord(n[1]) - ord('A') + 1, enumerate(column), 1) - 1

def cell_column(address: str):
    return column_id(re.match('(?P<column>[A-Z]+)\\d+', address).group('column'))

def rows_count(cells: dict[(int, int), int]):
    return max([k[0] for k in cells.keys()]) + 1

def columns_count(cells: dict[(int, int), int]):
    return max([k[1] for k in cells.keys()]) + 1

# Program
def execute_xltm(args: list[str]):
    if len(args) > 1:
        with ZipFile(sys.argv[1]) as zip:
            write_outputs(output_path(args), zip)
    else:
        sys.exit('Error Code 2: Input File Required')

def write_outputs(path: str, zip: ZipFile):
    write_images(path, zip)
    write_sheets(path, zip)

def output_path(args: list[str]):
    return args[2] if len(args) > 2 else './'

def is_executed():
    return __name__ == "__main__"

def read_file(path: str, zip: ZipFile):
    return ElementTree.parse(BytesIO(zip.read(path)))

def xml_elements(path: str, tree: ElementTree):
    return tree.findall(f'.//{{*}}{path}')

if is_executed():
    execute_xltm(sys.argv)