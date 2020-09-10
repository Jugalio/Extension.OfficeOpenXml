using DocumentFormat.OpenXml.Spreadsheet;
using Extension.Utilities.ClassExtensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Extension.OfficeOpenXml.Excel
{
    public class ExcelRow
    {
        /// <summary>
        /// The excel file this row is a part of
        /// </summary>
        public ExcelFile ExcelFile;

        /// <summary>
        /// The index of this row
        /// </summary>
        public uint RowIndex => ThisRow.RowIndex;

        /// <summary>
        /// The open xml row object
        /// </summary>
        public Row ThisRow;

        /// <summary>
        /// The cell wrapper objects on this row
        /// </summary>
        public List<ExcelCell> Cells = new List<ExcelCell>();

        /// <summary>
        /// Creates an empty row
        /// </summary>
        public ExcelRow(ExcelFile file, uint index)
        {
            ExcelFile = file;
            ThisRow = new Row();
            ThisRow.RowIndex = index;
        }

        /// <summary>
        /// Opens a new row element from an excel file
        /// </summary>
        public ExcelRow(ExcelFile file, Row row)
        {
            ExcelFile = file;
            ThisRow = row;

            int i = 1;
            foreach (Cell cell in ThisRow.ChildElements)
            {
                var nextIndex = cell.GetColumnIndex();
                for (int j = i; j < nextIndex; j++)
                {
                    Cells.Add(new ExcelCell(ExcelFile, string.Empty));
                }
                Cells.Add(new ExcelCell(ExcelFile, cell));
                i = nextIndex + 1;
            }
        }

        /// <summary>
        /// Adds a new cell to the row
        /// </summary>
        /// <param name="value"></param>
        public ExcelCell AddCell(string value)
        {
            var cell = new ExcelCell(ExcelFile, value);
            AddNewCell(cell);
            return cell;
        }

        /// <summary>
        /// Adds a new cell to the row and sets the cell reference
        /// </summary>
        /// <param name="cell"></param>
        private void AddNewCell(ExcelCell cell)
        {
            var column = Cells.LastOrDefault()?.ThisCell.GetNextColumnName() ?? "A";
            cell.ThisCell.CellReference = $"{column}{RowIndex}";
            ThisRow.Append(cell.ThisCell);
            Cells.Add(cell);
        }

        /// <summary>
        /// The value of a cell within the row by the index
        /// </summary>
        /// <returns></returns>
        public ExcelCell GetCellByColumnName(string name)
        {
            return Cells.FirstOrDefault(c => c.ThisCell.GetColumnName() == name);
        }

        /// <summary>
        /// The value of a cell within the row by the index
        /// </summary>
        /// <returns></returns>
        public ExcelCell GetCellByColumnIndex(uint index)
        {
            string start = "A";

            for (uint i = 1; i < index; i++)
            {
                start = start.IterateUpperLetter();
            }

            return Cells.FirstOrDefault(c => c.ThisCell.GetColumnName() == start);
        }

        public void CopyCellsFromOtherDocument(List<ExcelCell> cells)
        {
            cells.ForEach(c =>
            {
                var cell = new ExcelCell(ExcelFile, c);
                ThisRow.Append(cell.ThisCell);
                Cells.Add(cell);
            });
        }

    }
}
