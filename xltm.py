from pathlib import Path
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree
from xml.etree.ElementTree import Element
from functools import reduce
import sys
import re

def write_outputs(path: str, xl: ZipFile):
    write_images(path, xl)
    write_sheets(path, xl)

def write_images(path: str, xl: ZipFile):
    for (id, data, ext) in read_images(xl):
        Path(path).joinpath(Path(f'{id}{ext}')).write_bytes(data)

def write_sheets(path: str, xl: ZipFile):
    for (name, cells) in read_sheets(xl).items():
        Path(path).joinpath(Path(f'{name}.csv')).write_text(serialize_csv(cells))

def read_images(xl: ZipFile):
    return [(i, xl.read(p), Path(p).suffix) for p, i in images_info(xl)]

def read_sheets(xl: ZipFile) -> dict[str, list[list[str]]]:
    return {e.get('name') : cells_csv(read_cells(f'xl/worksheets/sheet{e.get('sheetId')}.xml', xl)) for e in xml_elements('sheet', read_file('xl/workbook.xml', xl))}

def image_id(element: Element):
    return int(element.get('Id')[3:])

def image_path(element: Element):
    return 'xl' + element.get('Target')[2:]

def image_elements(xl: ZipFile) -> list[Element]:
    return [e for e in xml_elements('Relationship', read_file('xl/richData/_rels/richValueRel.xml.rels', xl)) if e.get('Type').endswith('/image')]

def images_info(zip: ZipFile):
    return [(image_path(e), image_id(e)) for e in image_elements(zip)]

def cells_csv(cells: dict[(int, int), int]):
    return [[cells.get((r, c)) for c in range(columns_count(cells))] for r in range(rows_count(cells))]

def read_cells(sheetpath: str, xlfile: ZipFile) -> dict[str, list[list[str]]]:
    return {(cell_row(a := c.get('r')), cell_column(a)) : c.get('vm') for c in xml_elements('c', read_file(sheetpath, xlfile))}

def serialize_csv(cells: list[int]):
    return '\n'.join([','.join([c if c is not None else '' for c in r]) for r in cells])

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

def read_file(path: str, xl: ZipFile):
    return ElementTree.parse(BytesIO(xl.read(path)))

def xml_elements(path: str, tree: ElementTree):
    return tree.findall(f'.//{{*}}{path}')

def output_path(args: list[str]):
    return args[2] if len(args) > 2 else './'

if __name__ == "__main__":
    if len(args := sys.argv) > 1:
        with ZipFile(args[1]) as zip:
            write_outputs(output_path(args), zip)
    else:
        sys.exit('Error Code 2: Input File Required')