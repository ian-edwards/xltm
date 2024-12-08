from pathlib import Path
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree
from functools import reduce
import sys
import re

def read_images(xlfile: ZipFile):
    return {p[14:] : xlfile.read(p) for p in xlfile.namelist() if p.startswith('xl/media/image')}

def read_sheets(xlfile: ZipFile):
    return {e.get('name') + '.csv' : cells_matrix(read_cells(f'xl/worksheets/sheet{e.get('sheetId')}.xml', xlfile)) for e in read_xml('xl/workbook.xml', xlfile).findall(f'.//{{*}}sheet')}

def write_images(outdir: str, images: dict[str, bytes]):
    for (name, data) in images.items():
        Path(outdir).joinpath(name).write_bytes(data)

def write_sheets(outdir: str, sheets: dict[str, list[list[int]]]):
    for (name, cells) in sheets.items():
        Path(outdir).joinpath(name).write_text(serialize_csv(cells))

def cells_matrix(cells: dict[(int, int), int]) -> list[list[str]]:
    return [[cells.get((r, c)) for c in range(columns_count(cells))] for r in range(rows_count(cells))]

def read_cells(sheetpath: str, xlfile: ZipFile):
    return {(cell_row(a := c.get('r')), cell_column(a)) : int(c.get('vm')) for c in read_xml(sheetpath, xlfile).findall(f'.//{{*}}c')}

def serialize_csv(cells: list[int]):
    return '\n'.join([','.join([str(c) if c is not None else '' for c in r]) for r in cells])

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

def read_xml(path: str, xl: ZipFile):
    return ElementTree.parse(BytesIO(xl.read(path)))

if __name__ == "__main__":
    if (largs := len(args := sys.argv)) > 1:
        with ZipFile(args[1]) as xlfile:
            write_images(outpath := args[2] if largs > 2 else './', read_images(xlfile))
            write_sheets(outpath, read_sheets(xlfile))
    else:
        sys.exit('Error Code 2: Input File Required')