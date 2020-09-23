using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Extension.Utilities.ClassExtensions;
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

            foreach (Row row in SheetData.ChildElements)
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
        /// Moves a full column within this sheet
        /// </summary>
        /// <param name="currentColumnIndex"></param>
        /// <param name="newColumnIndex"></param>
        public void MoveColumn(uint currentColumnIndex, uint newColumnIndex)
        {
            Rows.ForEach(r => r.MoveColumn(currentColumnIndex, newColumnIndex));
            MoveColumnDef(currentColumnIndex, newColumnIndex);
        }

        /// <summary>
        /// Moves a full column within this sheet
        /// </summary>
        /// <param name="currentColumnId"></param>
        /// <param name="newColumnId"></param>
        public void MoveColumn(string currentColumnId, string newColumnId)
        {
            Rows.ForEach(r => r.MoveColumn(currentColumnId, newColumnId));
            MoveColumnDef(GetColumnIndex(currentColumnId), GetColumnIndex(newColumnId));
        }

        /// <summary>
        /// Moves the column def, which is its style definition
        /// </summary>
        internal void MoveColumnDef(uint currentColumnIndex, uint newColumnIndex)
        {
            var columns = _worksheet.GetFirstChild<Columns>().ChildElements.ToList().Cast<Column>();
            var target = columns.FirstOrDefault(c => c.Min <= currentColumnIndex && c.Max >= currentColumnIndex);
            if (target != null)
            {
                var styleToInsert = (Column)target.CloneNode(true);
                RemoveColumnStyle(currentColumnIndex);
                InsertColumnStyle(newColumnIndex, styleToInsert);
            }
        }

        /// <summary>
        /// Removes the full column
        /// </summary>
        /// <param name="columnId"></param>
        /// <param name="moveAllRightOf"></param>
        public void RemoveColumn(string columnId, bool moveAllRightOf = true)
        {
            var index = GetColumnIndex(columnId);
            Rows.ForEach(r => r.RemoveCell(columnId, moveAllRightOf));
            if (moveAllRightOf)
            {
                RemoveColumnStyle(index);
            }
        }

        /// <summary>
        /// Inserts the column style at a given index
        /// </summary>
        /// <param name="index"></param>
        internal void InsertColumnStyle(uint index, Column style)
        {
            style.Min = index;
            style.Max = index;
            var columns = _worksheet.GetFirstChild<Columns>().ChildElements.ToList().Cast<Column>();
            var target = columns.FirstOrDefault(c => c.Min <= index && c.Max >= index);
            if (target != null)
            {
                if (target.Min == target.Max || target.Min == index)
                {
                    //This also includes target
                    foreach (var c in columns.Where(c => c.Min >= index))
                    {
                        c.Min++;
                        c.Max++;
                    }
                    _worksheet.GetFirstChild<Columns>().InsertBefore(style, target);
                }
                else
                {
                    var splitMax = target.Max + 1;
                    target.Max = index - 1;
                    _worksheet.GetFirstChild<Columns>().InsertAfter(style, target);
                    var targetSplit = (Column)target.CloneNode(true);
                    targetSplit.Min = index + 1;
                    targetSplit.Max = splitMax;
                    _worksheet.GetFirstChild<Columns>().InsertAfter(targetSplit, style);

                    foreach (var c in columns.Where(c => c.Min > splitMax))
                    {
                        c.Min++;
                        c.Max++;
                    }
                }
            }
        }

        /// <summary>
        /// Removes the column style at a given index
        /// </summary>
        /// <param name="index"></param>
        internal void RemoveColumnStyle(uint index)
        {
            var count = _worksheet.GetFirstChild<Columns>().ChildElements.Count;
            var columns = _worksheet.GetFirstChild<Columns>().ChildElements.ToList().Cast<Column>();
            var target = columns.FirstOrDefault(c => c.Min <= index && c.Max >= index);
            if (target != null)
            {
                if (target.Min == target.Max)
                {
                    target.Remove();
                }
                else
                {
                    target.Max--;
                }

                foreach (var c in columns.Where(c => c.Min > index))
                {
                    c.Min--;
                    c.Max--;
                }
            }
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

        /// <summary>
        /// Get the columnid for the index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        private uint GetColumnIndex(string id)
        {
            uint index = 1;

            while (id != "A")
            {
                id = id.ReverseIterateUpperLetter();
                index++;
            }

            return index;
        }

    }
}
