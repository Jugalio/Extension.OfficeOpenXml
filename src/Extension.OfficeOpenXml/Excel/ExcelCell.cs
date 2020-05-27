using DocumentFormat.OpenXml.Spreadsheet;
using Extension.Utilities.ClassExtensions;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Text;

namespace Extension.OfficeOpenXml.Excel
{
    public class ExcelCell
    {
        /// <summary>
        /// The excel file this cell is a part of
        /// </summary>
        public ExcelFile ExcelFile;

        /// <summary>
        /// The open xml cell reference
        /// </summary>
        public Cell ThisCell;

        /// <summary>
        /// Creates a new cell object from open xml object
        /// </summary>
        /// <param name="cell"></param>
        public ExcelCell(ExcelFile file, Cell cell)
        {
            ExcelFile = file;
            ThisCell = cell;
        }

        /// <summary>
        /// Creates a new excel cell with the given value
        /// </summary>
        /// <param name="value"></param>
        public ExcelCell(ExcelFile file, string value)
        {
            ExcelFile = file;
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
        public ExcelCell(ExcelFile file, int value)
        {
            ExcelFile = file;
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

        /// <summary>
        /// Get the value of this cell, even if it is a shared string
        /// </summary>
        /// <returns></returns>
        public string GetValue()
        {
            string cellValue = string.Empty;

            if (ThisCell.DataType != null)
            {
                if (ThisCell.DataType == CellValues.SharedString)
                {
                    int id = -1;

                    if (int.TryParse(ThisCell.InnerText, out id))
                    {
                        SharedStringItem item = ExcelFile.GetSharedStringItemById(id);

                        if (item.Text != null)
                        {
                            cellValue = item.Text.Text;
                        }
                        else if (item.InnerText != null)
                        {
                            cellValue = item.InnerText;
                        }
                        else if (item.InnerXml != null)
                        {
                            cellValue = item.InnerXml;
                        }
                    }
                }
                else
                {
                    return ThisCell.InnerText;
                }
            }

            return cellValue;
        }

        /// <summary>
        /// Get the column name of this cell
        /// </summary>
        /// <returns></returns>
        public string GetColumnName()
        {
            var cellRef = ThisCell.CellReference?.Value;
            if (cellRef == null)
            {
                return string.Empty;
            }
            else
            {
                return GetCellNameFromCellRef(cellRef);
            }
        }

        /// <summary>
        /// Returns the column name of the previous column
        /// </summary>
        /// <returns></returns>
        public string GetPreviousColumnName()
        {
            return GetColumnName().ReverseIterateUpperLetter();
        }

        /// <summary>
        /// Returns the column name of the next column
        /// </summary>
        /// <returns></returns>
        public string GetNextColumnName()
        {
            return GetColumnName().IterateLowerLetter();
        }

        /// <summary>
        /// Get the column name from the cell reference string
        /// </summary>
        /// <param name="cellRef"></param>
        /// <returns></returns>
        private string GetCellNameFromCellRef(string cellRef)
        {
            var row = cellRef.GetNumericTail();
            return cellRef.Remove(cellRef.Length - row.Length);
        }

    }
}
