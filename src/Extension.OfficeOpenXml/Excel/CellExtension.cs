using DocumentFormat.OpenXml.Spreadsheet;
using Extension.Utilities.ClassExtensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Extension.OfficeOpenXml.Excel
{
    public static class CellExtension
    {
        /// <summary>
        /// Gets the column index of this cell 1 based as it is used in excel
        /// </summary>
        /// <returns></returns>
        public static int GetColumnIndex(this Cell ThisCell)
        {
            int columnNumber = 0;
            int mulitplier = 1;

            //working from the end of the letters take the ASCII code less 64 (so A = 1, B =2...etc)
            //then multiply that number by our multiplier (which starts at 1)
            //multiply our multiplier by 26 as there are 26 letters
            foreach (char c in ThisCell.GetColumnName().ToCharArray().Reverse())
            {
                columnNumber += mulitplier * ((int)c - 64);
                mulitplier = mulitplier * 26;
            }
            return columnNumber;
        }

        /// <summary>
        /// Get the column name of this cell
        /// </summary>
        /// <returns></returns>
        public static string GetColumnName(this Cell ThisCell)
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
        public static string GetPreviousColumnName(this Cell ThisCell)
        {
            return ThisCell.GetColumnName().ReverseIterateUpperLetter();
        }

        /// <summary>
        /// Returns the column name of the next column
        /// </summary>
        /// <returns></returns>
        public static string GetNextColumnName(this Cell ThisCell)
        {
            return ThisCell.GetColumnName().IterateLowerLetter();
        }

        /// <summary>
        /// Get the column name from the cell reference string
        /// </summary>
        /// <param name="cellRef"></param>
        /// <returns></returns>
        private static string GetCellNameFromCellRef(string cellRef)
        {
            var row = cellRef.GetNumericTail();
            return cellRef.Remove(cellRef.Length - row.Length);
        }

    }
}
