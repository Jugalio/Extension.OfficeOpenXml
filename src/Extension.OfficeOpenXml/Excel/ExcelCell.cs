using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Text;

namespace Extension.OfficeOpenXml.Excel
{
    public class ExcelCell
    {
        public Cell ThisCell;

        /// <summary>
        /// Creates a new excel cell with the given value
        /// </summary>
        /// <param name="value"></param>
        public ExcelCell(string value)
        {
            ThisCell = new Cell()
            {
                CellValue = new CellValue(value),
                DataType = CellValues.String,
            };
        }

        /// <summary>
        /// Creates a new excel cell with the given value
        /// </summary>
        /// <param name="value"></param>
        public ExcelCell(int value)
        {
            ThisCell = new Cell()
            {
                CellValue = new CellValue(value.ToString()),
                DataType = CellValues.Number,
            };
        }

        /// <summary>
        /// Adds a style to the cell
        /// </summary>
        /// <param name="styleIndex"></param>
        public void AddStyle(uint styleIndex)
        {
            ThisCell.StyleIndex = styleIndex;
        }

    }
}
