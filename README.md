# ExcelHelper

a wrapper with a few convenience shortcuts for [ClosedXml](https://github.com/ClosedXML/ClosedXML)

Features:
- Automatic total row with formula
- Automatic Pivot Table
- Shortcut method for formulas

## What is ClosedXml?
from their own repository description:
> ClosedXML is a .NET library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files. 
> It aims to provide an intuitive and user-friendly interface to dealing with the underlying OpenXML API.

## How to use ExcelHelper

outputs a table with a total row. Only numeric columns are taken into consideration
```
var SomeDataTable = GetDataTable(sql);
            
var file = AppDomain.CurrentDomain.BaseDirectory + "/output/rowtotal.xlsx";
            
ExcelHelper eh = ExcelHelper.Create(file)
                        .Dump(SomeDataTable, "SHEET1")
                        .AddRowTotal("SHEET1", "A", "Tot.", true)
                        .Save();
eh.Dispose();
```
outputs a pivot table
```
string sql = "SELECT * FROM Invoices ORDER BY OrderDate";
var results = GetDataTable(config, sql, null);
var file = AppDomain.CurrentDomain.BaseDirectory + "/output/pivot_at_once.xlsx";
           
ExcelHelper eh = ExcelHelper.Create(file)
                        .Dump(results, "SHEET1")
                        .DoPivotTable("SHEET1",
                                      "SHEET2",
                                      new string[] { "ShipCountry" },
                                      new string[] { "ProductName" },
                                      new string[] { "ExtendedPrice" }
                                     )
                        .Save();
eh.Dispose();
```
## Final Notes
[ClosedXml](https://github.com/ClosedXML/ClosedXML) has not reached a stable 1.0 release yet, so api may change.
I'm using my wrapper in production and it happened to me a few times to make tweaks.

All my wrapper does (and more) can be done by [ClosedXml](https://github.com/ClosedXML/ClosedXML)  companion class [ClosedXml.Report](https://github.com/ClosedXML/ClosedXML.Report). 
The only drawback, at least for me atm, is that you need to use templates which may or maynot suit your needs.

 [ClosedXml.Report](https://github.com/ClosedXML/ClosedXML.Report) looks pretty promising anyway and I wish to look into it soon. From their very own description:
> ClosedXML.Report is a tool for report generation and data analysis 
> in .NET applications through the use of Microsoft Excel. 
> It is a .NET-library for report generation Microsoft Excel without requiring Excel to be installed on the machine that's running the code. 
> With ClosedXML.Report, you can easily export any data from your .NET classes to Excel using the XLSX-template.
