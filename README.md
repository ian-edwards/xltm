# xltm

xltm is a Python library that converts Excel files into csv tilemaps with corresponding image files

Excel Desktop and Excel Web can be used to generate tilemaps by inserting [pictures in-cell](https://support.microsoft.com/en-us/office/insert-picture-in-cell-in-excel-e9317aee-4294-49a3-875c-9dd95845bab0)

This can be a quick and easy way to generate a large number of visual tilemaps

Unfortunately, Excel files are not a convenient format to use in other projects

xltm simplifies this by converting complex Excel files into easier to use simple csv and image files

## Images

All images are extracted from the workbook and are written to the output directory named with unique ids
```
1.png
2.png
3.jpeg
5.jpg
```
Ids are integers that may not be contiguous

Extensions depend on the source format of the image in Excel

## Sheets

Each sheet in the workbook is written to the output directory as a csv file named the same as the sheet
```
forest1.csv
forest2.csv
```

Each entry in the csv file corresponds to an image id
```
3,3,3,3,3
3,1,1,2,3
3,2,2,5,3
1,1,1,1,2
```

Empty csv entries correspond to tiles that have no images
```
3,3,3,,3
3,,1,2,3
3,2,,5,3
,,,,
```
## Usage

xltm may be invoked directly via the command line with either 1 or 2 input parameters

```console
python xltm.py ./input/file/path.xlsx
```

```console
python xltm.py ./input/file/path.xlsx ./output/directory/path
```

The first parameter is required and is the path of the Excel file to process

The second parameter is optional and is the directory where outputs will be written

If the second parameter is omitted outputs are written to the current directory

xltm may also be invoked in code

```python
import xltm
from zipfile import ZipFile

with ZipFile('./input/file/path.xlsx') as xlfile:
    images = xltm.read_images(xlfile) # dict[str, data]
    sheets = xltm.read_sheets(xlfile) # dict[str, list[list[int]]]
    xltm.write_images('./output/directory/path/', images)
    xltm.write_sheets('./output/directory/path/', sheets)
```

## Compatibility

xltm is compatible with Pyton >= 3.8

xltm is compatible with xlsx and xlsm files

xltm is not compatible with xlsb, xls or any other Excel files

## Authors

[Ian Edwards](mailto:ian.contact@proton.me)

## License

This project is licensed under the [MIT License](https://opensource.org/license/MIT):

Copyright (c) 2024 Ian Edwards

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.