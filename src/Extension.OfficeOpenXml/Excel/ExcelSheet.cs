using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Extension.OfficeOpenXml.Excel
{
    public class ExcelSheet
    {
        /// <summary>
        /// The name of this sheet
        /// </summary>
        public string Name => ThisSheet.Name;

        /// <summary>
        /// The excel file this sheet is a part of
        /// </summary>
        public ExcelFile ExcelFile;

        /// <summary>
        /// The worksheetpart for this sheet
        /// </summary>
        public WorksheetPart WorksheetPart;

        /// <summary>
        /// The open xml sheet instance
        /// </summary>
        public Sheet ThisSheet;

        /// <summary>
        /// The worksheet element which belongs to this sheet
        /// </summary>
        private Worksheet _worksheet;

        /// <summary>
        /// The data of this sheet
        /// </summary>
        public SheetData SheetData;

        /// <summary>
        /// The rows in this excel sheet
        /// </summary>
        public List<ExcelRow> Rows = new List<ExcelRow>();

        /// <summary>
        /// Creates an empty sheet
        /// </summary>
        /// <param name="file"></param>
        /// <param name="name"></param>
        public ExcelSheet(ExcelFile file, string name, uint sheetId)
        {
            ExcelFile = file;

            //Add worksheetpart
            WorksheetPart = ExcelFile.WorkbookPart.AddNewPart<WorksheetPart>();
            SheetData = new SheetData();
            _worksheet = new Worksheet(SheetData);
            WorksheetPart.Worksheet = _worksheet;

            var id = ExcelFile.WorkbookPart.GetIdOfPart(WorksheetPart);
            CreateEmptySheet(name, id, sheetId);
        }

        /// <summary>
        /// Creates an empty sheet
        /// </summary>
        /// <param name="file"></param>
        /// <param name="name"></param>
        public ExcelSheet(ExcelFile file, string name, uint sheetId, ExcelSheet referenceSheet)
        {
            ExcelFile = file;

            //Add worksheetpart
            WorksheetPart = ExcelFile.WorkbookPart.AddNewPart<WorksheetPart>();
            SheetData = new SheetData();
            _worksheet = new Worksheet(SheetData);
            WorksheetPart.Worksheet = _worksheet;

            var id = ExcelFile.WorkbookPart.GetIdOfPart(WorksheetPart);
            ThisSheet = (Sheet)referenceSheet.ThisSheet.CloneNode(false);
            ThisSheet.Id = id;
            ThisSheet.Name = name;
            ThisSheet.SheetId = sheetId;

            var columns = new Columns();
            columns.Append(referenceSheet._worksheet.GetFirstChild<Columns>().ChildElements.Select(e => e.CloneNode(false)));
            _worksheet.InsertAt(columns, 0);
            ExcelFile.Sheets.Append(ThisSheet);
        }

        /// <summary>
        /// Creates a new sheet wrapper object from a 
        /// open xml sheet class object
        /// </summary>
        /// <param name="sheet"></param>
        public ExcelSheet(ExcelFile file, Sheet sheet)
        {
            ExcelFile = file; 
            ThisSheet = sheet;

            WorksheetPart = (WorksheetPart)ExcelFile.WorkbookPart.GetPartById(sheet.Id);
            _worksheet = WorksheetPart.Worksheet;
            SheetData = _worksheet.GetFirstChild<SheetData>();

            foreach(Row row in SheetData.ChildElements)
            {
                Rows.Add(new ExcelRow(ExcelFile, row));
            }
        }

        /// <summary>
        /// Adds a new row to the sheet
        /// </summary>
        /// <returns></returns>
        public ExcelRow AppendNewRow()
        {
            var index = Rows.Count == 0 ? 1 : Rows.Select(r => r.RowIndex).Max() + 1;
            var row = new ExcelRow(ExcelFile, index);
            SheetData.AppendChild(row.ThisRow);
            Rows.Add(row);
            return row;
        }

        /// <summary>
        /// Adds a new row to the sheet
        /// </summary>
        /// <returns></returns>
        public ExcelRow AppendRow(ExcelRow referenceRow, bool clone = false)
        {
            var index = Rows.Count == 0 ? 1 : Rows.Select(r => r.RowIndex).Max() + 1;
            var node = (Row)referenceRow.ThisRow.CloneNode(false);
            var row = new ExcelRow(ExcelFile, node);
            row.ThisRow.RowIndex = index;
            row.InsertCells(referenceRow.Cells, clone);
            SheetData.AppendChild(row.ThisRow);
            Rows.Add(row);
            return row;
        }

        /// <summary>
        /// Adds a new row to the sheet
        /// </summary>
        /// <returns></returns>
        public ExcelRow InsertRowAt(ExcelRow referenceRow, uint index, bool clone = false)
        {
            var occupied = Rows.FirstOrDefault(r => r.RowIndex == index);
            if (occupied != null)
            {
                Rows.Where(r => r.RowIndex > index).ToList().ForEach(r => r.ThisRow.RowIndex++);
                occupied.ThisRow.RowIndex++;
            }

            var node = (Row)referenceRow.ThisRow.CloneNode(false);
            var row = new ExcelRow(ExcelFile, node);
            row.ThisRow.RowIndex = index;
            row.InsertCells(referenceRow.Cells, clone);
            SheetData.AppendChild(row.ThisRow);
            Rows.Add(row);
            return row;
        }

        /// <summary>
        /// Creates an empty sheet and attach it to the excel file
        /// </summary>
        /// <param name="name"></param>
        private void CreateEmptySheet(string name, string id, uint sheetId)
        {
            ThisSheet = new Sheet()
            {
                Id = id,
                Name = name,
                SheetId = sheetId,
            };
            ExcelFile.Sheets.Append(ThisSheet);
        }

    }
}
