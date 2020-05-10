using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace Extension.OfficeOpenXml.Excel
{
    public class ExcelRow
    {

        public Row ThisRow;
        public List<ExcelCell> Cells = new List<ExcelCell>();

        /// <summary>
        /// Creates an empty row
        /// </summary>
        public ExcelRow()
        {
            ThisRow = new Row();
        }

        /// <summary>
        /// Adds a new cell to the row
        /// </summary>
        /// <param name="value"></param>
        public void AddCell(string value)
        {
            var cell = new ExcelCell(value);
            ThisRow.Append(cell.ThisCell);
            Cells.Add(cell);
        }

        /// <summary>
        /// Adds a new cell to the row
        /// </summary>
        /// <param name="value"></param>
        public void AddCell(int value)
        {
            var cell = new ExcelCell(value);
            ThisRow.Append(cell.ThisCell);
            Cells.Add(cell);
        }

    }
}
