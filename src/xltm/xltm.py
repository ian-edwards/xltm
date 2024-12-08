from pathlib import Path
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree
from xml.etree.ElementTree import Element
from dataclasses import dataclass
from functools import reduce
import sys
import re

@dataclass
class FileImage:
    id: int
    extension: str
    data: bytes

@dataclass
class FileSheet:
    name: str
    cells: list[list[str]]

def write_outputs(outdir: str, xlfile: ZipFile):
    write_images(outdir, read_images(xlfile))
    write_sheets(outdir, read_sheets(xlfile))

def write_images(outdir: str, images: list[FileImage]):
    for image in images:
        write_image(outdir, image)

def write_sheets(outdir: str, sheets: list[FileSheet]):
    for sheet in sheets:
        write_sheet(outdir, sheet)

def write_sheet(outdir: str, sheet: FileSheet):
    Path(outdir).joinpath(Path(f'{sheet.name}.csv')).write_text(serialize_csv(sheet.cells))

def write_image(outdir: str, image: FileImage):
    Path(outdir).joinpath(Path(f'{image.id}{image.extension}')).write_bytes(image.data)

def read_images(xlfile: ZipFile):
    return [FileImage(i, Path(p).suffix, xlfile.read(p)) for p, i in images_info(xlfile)]

def read_sheets(xlfile: ZipFile):
    return [FileSheet(e.get('name'), cells_matrix(read_cells(f'xl/worksheets/sheet{e.get('sheetId')}.xml', xlfile))) for e in xml_elements('sheet', read_file('xl/workbook.xml', xlfile))]

def image_id(element: Element):
    return int(element.get('Id')[3:])

def image_path(element: Element):
    return 'xl' + element.get('Target')[2:]

def image_elements(xlfile: ZipFile) -> list[Element]:
    return [e for e in xml_elements('Relationship', read_file('xl/richData/_rels/richValueRel.xml.rels', xlfile)) if e.get('Type').endswith('/image')]

def images_info(xlfile: ZipFile):
    return [(image_path(e), image_id(e)) for e in image_elements(xlfile)]

def cells_matrix(cells: dict[(int, int), int]) -> list[list[str]]:
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
        with ZipFile(args[1]) as xlfile:
            write_outputs(output_path(args), xlfile)
    else:
        sys.exit('Error Code 2: Input File Required')