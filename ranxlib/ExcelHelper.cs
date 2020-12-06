using ClosedXML.Excel;
using System;
using System.IO;

namespace ranxlib
{
    public class ExcelHelper : IDisposable
    {
        protected FileInfo _inputFile = null;
        protected XLWorkbook _workbook;
        protected string _defaultSummarySheetName = "SUMMARY";
        protected SaveOptions _defaultSaveOptions = null;

        /// <summary>
        ///         ''' Apre file "inputFile"
        ///         ''' </summary>
        ///         ''' <param name="inputFile">percorso a file excel</param>
        ///         ''' <returns></returns>
        public static ExcelHelper Create(string inputFile, bool mustExists = false, string summarySheetName = "SUMMARY")
        {
            return new ExcelHelper(inputFile, mustExists, summarySheetName);
        }
        protected void _CreateFile(string inputFile)
        {
            this._workbook = new XLWorkbook(XLEventTracking.Disabled);

            var worksheet = this._CreateSummary();

            _workbook.SaveAs(inputFile);
            _workbook.Dispose();
        }
        protected IXLWorksheet _CreateSummary()
        {
            var worksheet = _workbook.Worksheets.Add(this._defaultSummarySheetName);
            worksheet.ActiveCell = worksheet.Cell(1, 10);
            worksheet.ActiveCell.Value = this._defaultSummarySheetName;
            worksheet.ActiveCell.Style.Font.Bold = true;
            return worksheet;
        }
        public ExcelHelper RemoveSummary()
        {
            this.RemoveSheet(this._defaultSummarySheetName);
            return this;
        }
        protected void RemoveSheet(string sheetName)
        {
            try
            {
                IXLWorksheet sheet = this._workbook.Worksheet(sheetName);
                if (sheet != null)
                {
                    _workbook.Worksheets.Delete(sheet.Name);
                }

            }
            catch (Exception ex)
            {
                //SD.Log(ex.Message, SD.LogLevel.Warning, ex.StackTrace);
            }
        }
        protected ExcelHelper(string inputFile, bool mustExists = false, string summarySheetName = "SUMMARY")
        {

            this._defaultSummarySheetName = summarySheetName;
            this._defaultSaveOptions = new SaveOptions();
            _defaultSaveOptions.EvaluateFormulasBeforeSaving = false;
            _defaultSaveOptions.ValidatePackage = false;
            bool fileExists = File.Exists(inputFile);
            if (fileExists == false & mustExists == true)
                throw new Exception("File " + inputFile + " not Found");
            else if (fileExists == false & mustExists == false)
            {
                try
                {
                    this._CreateFile(inputFile);
                    this._workbook = new XLWorkbook(XLEventTracking.Disabled);
                }
                catch (Exception ex)
                {

                    //SD.Log(ex.Message + " " + ex.Source, SD.LogLevel.Error, ex.StackTrace);
                    return;
                }
            }
            else if (fileExists == true)
            {
                try
                {
                    this._workbook = new XLWorkbook(XLEventTracking.Disabled);
                }
                catch (Exception ex)
                {
                    //SD.Log(ex.Message, SD.LogLevel.Error, ex.StackTrace);
                }
            }
            this._inputFile = new FileInfo(inputFile);
        }
        /// <summary>
        ///         ''' Le datatable devono avere .Name=nome del worksheet per questo controllo prima se esiste già e lo cancello
        ///         ''' </summary>
        ///         ''' <param name="records"></param>
        ///         ''' <param name="workSheetName"></param>
        ///         ''' <returns></returns>
        public ExcelHelper Dump(System.Data.DataTable records, string workSheetName, bool forceWSheetRecreate = false)
        {
            this.RemoveSheet(workSheetName);
            this._workbook.Worksheets.Add(records, workSheetName);

            return this;
        }
        /// <summary>
        ///         ''' can't touch already created pivot-table
        ///         ''' https://github.com/ClosedXML/ClosedXML/pull/124
        ///         ''' </summary>
        ///         ''' <param name="sheetSource"></param>
        ///         ''' <param name="sheetDest"></param>
        ///         ''' <param name="rows"></param>
        ///         ''' <param name="cols"></param>
        ///         ''' <returns></returns>
        public ExcelHelper DoPivotTable(string sheetSource, string sheetDest, string[] rows = null, string[] cols = null, string[] datafields = null)
        {
            try
            {
                IXLWorksheet worksheetSource = null;
                IXLTable sourceTable = null;

                if (_workbook.Worksheets.Count < 1)
                {
                    _workbook = new XLWorkbook(this._inputFile.FullName);
                }
                worksheetSource = _workbook.Worksheets.Worksheet(sheetSource);
                sourceTable = worksheetSource.Table(0);
                // TODO PRIMA CONTROLLARE SE ESISTE E nel caso cancella
                this.RemoveSheet(sheetDest); // <<TODO invece di remove che in questo caso crea un errore cercare se è possibile chiamare REFRESH
                IXLWorksheet pivotTableSheet = _workbook.Worksheets.Add(sheetDest);

                IXLPivotTable pivoTable = pivotTableSheet.PivotTables.Add("PivotTable", pivotTableSheet.Cell(1, 1), sourceTable.AsRange());

                foreach (string r in rows)
                {
                    if (r.Trim() != "")
                        pivoTable.RowLabels.Add(r);
                }
                foreach (string c in cols)
                {
                    if (c.Trim() != "")
                        pivoTable.ColumnLabels.Add(c);
                }
                foreach (string d in datafields)
                {
                    if (d.Trim() != "")
                        pivoTable.Values.Add(d);
                }
            }
            // i filtri non sono al momento supportati https://github.com/ClosedXML/ClosedXML/issues/218

            catch (Exception ex)
            {
                //SD.Log(ex.Message, SD.LogLevel.Error, ex.StackTrace);

                return this;
            }
            return this;
        }
        public ExcelHelper AddSheet(string sheetName)
        {
            try
            {
                if (this._workbook == null) { 
                    this._workbook.AddWorksheet(sheetName);
                }
                else {
                    // SD.Log("No workbook open...aborting", SD.LogLevel.Info);
                }
            }
            catch (Exception ex)
            {
               // SD.Log(sheetName + " already exists", SD.LogLevel.Warning, ex.StackTrace);
            }


            return this;
        }
        /// <summary>
        ///         ''' TOD: studiare https://github.com/ClosedXML/ClosedXML/wiki/Using-Formulas
        ///         ''' </summary>
        ///         ''' <param name="sheetName"></param>
        ///         ''' <param name="cellAddress"></param>
        ///         ''' <param name="formula"></param>
        ///         ''' <returns></returns>
        public ExcelHelper AddFormula(string sheetName, IXLAddress cellAddress, string formula)
        {
            try
            {
                IXLWorksheet ActiveWorkSheet = this._workbook.Worksheets.Worksheet(sheetName);
                IXLCell ActiveCell = ActiveWorkSheet.Cell(cellAddress);
                ActiveCell.FormulaA1 = formula;
            }
            // _workbook.CalculateMode = autoCalc
            catch (Exception ex)
            {
               // SD.Log(ex.Message, SD.LogLevel.Error, ex.StackTrace);
            }


            return this;
        }
        /// <summary>
        ///         ''' TODO AGGIUNGERE PARAMETRO ROUNDING OPPURE CELLFORMAT
        ///         ''' </summary>
        ///         ''' <param name="sheetName"></param>
        ///         ''' <param name="labelCell"></param>
        ///         ''' <param name="labelCellValue"></param>
        ///         ''' <param name="excludeHeaders"></param>
        ///         ''' <returns></returns>
        public ExcelHelper AddRowTotal(string sheetName, string labelCell = "A", string labelCellValue = "TOT.", bool excludeHeaders = true)
        {
            IXLRange range = this.GetRangeUsed(sheetName);
            if (range == null)
            {
                //SD.Log("range not found in " + sheetName, SD.LogLevel.Error);
                return this;
            }
            int rowsNumber = range.RowCount();
            IXLCell lastCellUsed = range.LastCellUsed();
            IXLColumn lastColUsed = lastCellUsed.WorksheetColumn();
            IXLRow lastRowUsed = lastCellUsed.WorksheetRow();
            string lastColLetter = lastColUsed.ColumnLetter();
            int lastRowNumber = lastRowUsed.RowNumber();
            IXLRows rows = lastRowUsed.InsertRowsBelow(1);
            //IXLRow newRow = rows.Last();
            var ws = this._workbook.Worksheets.Worksheet(sheetName);
            IXLRow newRow = ws.LastRowUsed().RowBelow();

            if (labelCell.Trim() != "")
            {
                newRow.Cell(labelCell).Value = labelCellValue;
                newRow.Cell(labelCell).Style.Font.Bold = true;
            }
            var firstTotalCellAddress = newRow.FirstCell().CellRight().Address;
            var lastTotalCellAddress = newRow.Cell(lastColLetter).Address;
            IXLRange rangeTotal = this.GetRangeUsed(sheetName, firstTotalCellAddress, lastTotalCellAddress);
            //int i = rangeTotal.Cells().Count() + 1;
            int i = rangeTotal.ColumnCount() + 1;
            int firstDataRowIndex = 0;
            // escludo la riga delle intestazioni
            if (excludeHeaders)
            {
                firstDataRowIndex = 2;
            }
                
            for (int k = 1; k <= i; k++)
            {
                XLDataType colDataType = newRow.Cell(k).CellAbove(1).DataType;
                if (colDataType == XLDataType.Number)
                {
                    string colLetter = newRow.Cell(k).Address.ColumnLetter;
                    string formula = "=SUM(" + colLetter + firstDataRowIndex.ToString() + ":" + colLetter + rowsNumber.ToString() + ")";
                    this.AddFormula(sheetName, newRow.Cell(k).Address, formula);
                }
            }
            newRow.AsRange().RangeUsed().Style.Border.TopBorder = XLBorderStyleValues.Thick;
            return this;
        }
        public IXLRange GetRangeUsed(string sheetName, IXLAddress startAddress, IXLAddress endAddress)
        {
            IXLRange range = null/* TODO Change to default(_) if this is not a reference type */;
            if (_workbook.Worksheets.Count < 1)
                _workbook = new XLWorkbook(this._inputFile.FullName);
            try
            {
                var sheet = this._workbook.Worksheets.Worksheet(sheetName);
                range = this._workbook.Worksheets.Worksheet(sheetName).Range(startAddress, endAddress);
            }
            catch (Exception ex)
            {
                //SD.Log(ex.Message, SD.LogLevel.Error, ex.StackTrace);
            }

            return range;
        }
        public IXLRange GetRangeUsed(string sheetName)
        {
            IXLRange range = null;
            if (_workbook.Worksheets.Count < 1)
                _workbook = new XLWorkbook(this._inputFile.FullName);
            try
            {
                range = this._workbook.Worksheets.Worksheet(sheetName).RangeUsed();
            }
            catch (Exception ex)
            {
                //SD.Log("sheet " + sheetName + " not found", SD.LogLevel.Error, ex.StackTrace);
            }

            return range;
        }
        public ExcelHelper Save(SaveOptions options = null)
        {
            try
            {
                if (options == null)
                {
                    options = this._defaultSaveOptions;
                }
                this._workbook.SaveAs(this._inputFile.FullName, options);
            }
            catch (Exception ex)
            {
                //SD.Log(ex.Message, SD.LogLevel.Error, ex.StackTrace);
                throw;
            }

            return this;
        }
        public void SaveAsStream(ref Stream stream, SaveOptions options = null)
        {
            try
            {
                if (options == null)
                    options = this._defaultSaveOptions;
                this._workbook.SaveAs(stream, options);
            }
            catch (Exception ex)
            {
                //SD.Log(ex.Message, SD.LogLevel.Error, ex.StackTrace);
                throw;
            }
        }

        private bool disposedValue; // Per rilevare chiamate ridondanti

        // IDisposable
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                    // TODO: eliminare lo stato gestito (oggetti gestiti).
                    this._workbook.Dispose();
            }
            disposedValue = true;
        }


        // Questo codice viene aggiunto da c# per implementare in modo corretto il criterio Disposable.
        public void Dispose()
        {
            // Non modificare questo codice. Inserire sopra il codice di pulizia in Dispose(disposing As Boolean).
            Dispose(true);
        }
    }
}
