# ExcelHelper

a wrapper with a few convenience shortcuts for [ClosedXml](https://github.com/ClosedXML/ClosedXML)

Features:
- Automatic total row with formula
- Automatic Pivot Table
- Shortcut method for formulas

## What is ClosedXml?
> ClosedXML is a .NET library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files. 
> It aims to provide an intuitive and user-friendly interface to dealing with the underlying OpenXML API.

## How to use ExcelHelper

output a table with a total row. Only numeric columns are taken into consideration
```
            var SomeDataTable = GetDataTable(sql);
            
            var file = AppDomain.CurrentDomain.BaseDirectory + "/output/rowtotal.xlsx";
            if (System.IO.File.Exists(file))
            {
                System.IO.File.Delete(file);
            }
            ExcelHelper eh = ExcelHelper.Create(file)
                        .Dump(results, "SHEET1")
                        .AddRowTotal("SHEET1", "A", "Tot.", true)
                        .Save();
            eh.Dispose();
```


