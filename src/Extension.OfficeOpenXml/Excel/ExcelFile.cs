using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Extension.OfficeOpenXml.Excel
{
    public class ExcelFile
    {
        public SpreadsheetDocument Document;

        private WorkbookPart _workbookPart;
        public Workbook Workbook;

        private WorksheetPart _worksheetPart;
        private Worksheet _worksheet;
        public SheetData SheetData;
        public Sheets Sheets;

        private WorkbookStylesPart _workbookStylesPart;
        public Stylesheet Stylesheet;

        public List<ExcelSheet> SheetList = new List<ExcelSheet>();

        /// <summary>
        /// Finalizer in order to clean
        /// </summary>
        ~ExcelFile()
        {
            if (Document != null) Document.Close();
        }

        /// <summary>
        /// Opens a new excel file
        /// </summary>
        /// <param name="fileName"></param>
        public void Open(string fileName, bool editable)
        {
            Document = SpreadsheetDocument.Open(fileName, editable);
            LoadDocument();
        }

        /// <summary>
        /// Creates a new excel file and returns an instance for it
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public void Create(string fileName)
        {
            Document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);

            GenerateEmptyChildElements();
            AddSheet("Sheet1");
        }

        /// <summary>
        /// Adds a new sheet
        /// </summary>
        /// <param name="name"></param>
        public void AddSheet(string name)
        {
            var maxId = SheetList.Select(s => s.ThisSheet.SheetId).Max() ?? 0;

            var id = _workbookPart.GetIdOfPart(_worksheetPart);
            var sheet = new ExcelSheet();
            sheet.Create(this, name, id, maxId + 1);
            SheetList.Add(sheet);
        }

        /// <summary>
        /// Loads an excel document
        /// </summary>
        private void LoadDocument()
        {
            //Add workbook part
            _workbookPart = Document.WorkbookPart;
            Workbook = _workbookPart.Workbook;

            //Add worksheetpart
            _worksheetPart = _workbookPart.WorksheetParts.FirstOrDefault();
            SheetData = _worksheet.GetFirstChild<SheetData>();
            _worksheet = _worksheetPart.Worksheet;

            //Adds an empty sheets section
            Sheets = Workbook.Sheets;

            //Adds an empty stylesheet
            _workbookStylesPart = _workbookPart.WorkbookStylesPart;
            Stylesheet = _workbookStylesPart.Stylesheet;
        }

        /// <summary>
        /// Generates the workbook parts with an empty workbook
        /// </summary>
        private void GenerateEmptyChildElements()
        {
            //Add workbook part
            _workbookPart = Document.AddWorkbookPart();
            Workbook = new Workbook();
            _workbookPart.Workbook = Workbook;

            //Add worksheetpart
            _worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            SheetData = new SheetData();
            _worksheet = new Worksheet(SheetData);
            _worksheetPart.Worksheet = _worksheet;

            //Adds an empty sheets section
            Sheets = new Sheets();
            Workbook.AppendChild(Sheets);

            //Adds an empty stylesheet
            _workbookStylesPart = _workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet = CreateDefaultStyleSheet();
            _workbookStylesPart.Stylesheet = Stylesheet;
        }

        /// <summary>
        /// Creates a default style sheet
        /// </summary>
        /// <returns></returns>
        private Stylesheet CreateDefaultStyleSheet()
        {
            Stylesheet styleSheet = null;

            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 }
                ));

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }) // Index 0 - default
                );

            Borders borders = new Borders(
                    new Border() // Index 0 - default
                );

            CellFormats cellFormats = new CellFormats(
                    new CellFormat() // Index 0 - default
                );

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }

        /// <summary>
        /// Saves the workbook and the document
        /// </summary>
        public void Save()
        {
            Workbook.Save();
            Document.Save();
        }
    }
}
