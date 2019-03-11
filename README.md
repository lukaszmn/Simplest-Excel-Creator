# Simplest Excel Creator

This project is a very simple and robust OpenXML file creator - it allows creating an *xlsx* Excel file.

## Limitations

It provides no means of formatting and currently allows providing only the following types of values to cells:
* numbers
* percentages (without formatting of decimal places - it will be shown rounded, but the stored value will have the original precision)
* text

Should you require other features, feel free to extend this library or find a more complex library - there are many available.

## Usage

The usage is very simple: `ExcelGenerator.Generate(data, path)`, where *data* stores two-dimensional information about values for cells
and *path* is the full path to the XLSX file that should be written.

## Reference

Some documentation on the Open XML format can be found at

https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet?view=openxml-2.8.1
