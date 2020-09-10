using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Extension.Utilities.ClassExtensions;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
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
        /// The datatype of this cell
        /// </summary>
        public CellValues DataType => ThisCell.DataType ?? CellValues.String;

        /// <summary>
        /// Gets the value of this cell as a string
        /// </summary>
        public string Value
        {
            get
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
                else if (ThisCell.CellValue != null)
                {
                    cellValue = ThisCell.CellValue.InnerText;
                }

                return cellValue;
            }
        }

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
        public ExcelCell(ExcelFile file, string value, uint styleIndex, CellValues dataType)
        {
            ExcelFile = file;

            if (dataType == CellValues.SharedString)
            {
                var index = SetCellValue(value);
                ThisCell = new Cell()
                {

                    DataType = new EnumValue<CellValues>(CellValues.SharedString),
                    CellValue = new CellValue(index.ToString()),
                    StyleIndex = styleIndex,
                };
            }
            else
            {
                ThisCell = new Cell()
                {
                    CellValue = new CellValue(value),
                    StyleIndex = styleIndex,
                };
            }
        }

        /// <summary>
        /// Creates a new excel from a given cell
        /// </summary>
        /// <param name="value"></param>
        public ExcelCell(ExcelFile file, ExcelCell refCell)
        {
            ExcelFile = file;

            if (refCell.ThisCell.DataType == null)
            {
                var index = SetCellValue(refCell.Value);
                ThisCell = new Cell()
                {
                    CellValue = new CellValue(refCell.Value),
                };
            }
            else if (refCell.ThisCell.CellFormula != null)
            {

            }
            else if (refCell.ThisCell.DataType == CellValues.SharedString)
            {
                var index = SetCellValue(refCell.Value);
                ThisCell = new Cell()
                {
                    DataType = new EnumValue<CellValues>(CellValues.SharedString),
                    CellValue = new CellValue(index.ToString()),
                };
            }
            else
            {
                ThisCell = new Cell()
                {
                    CellValue = new CellValue(refCell.Value),
                };
            }

            ThisCell.StyleIndex = refCell.ThisCell.StyleIndex;
            ThisCell.CellFormula = refCell.ThisCell.CellFormula != null ? new CellFormula(refCell.ThisCell.CellFormula.InnerText) : null;
        }

        /// <summary>
        /// Creates a new excel cell with the given value
        /// </summary>
        /// <param name="value"></param>
        public ExcelCell(ExcelFile file, string value, CellValues dataType)
        {
            ExcelFile = file;

            if (dataType == CellValues.SharedString)
            {
                var index = SetCellValue(value);
                ThisCell = new Cell()
                {

                    DataType = new EnumValue<CellValues>(CellValues.SharedString),
                    CellValue = new CellValue(index.ToString()),
                };
            }
            else
            {
                ThisCell = new Cell()
                {
                    CellValue = new CellValue(value),
                };
            }
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
        /// Gets the column index of this cell 1 based as it is used in excel
        /// </summary>
        /// <returns></returns>
        public int GetColumnIndex()
        {
            int columnNumber = 0;
            int mulitplier = 1;

            //working from the end of the letters take the ASCII code less 64 (so A = 1, B =2...etc)
            //then multiply that number by our multiplier (which starts at 1)
            //multiply our multiplier by 26 as there are 26 letters
            foreach (char c in GetColumnName().ToCharArray().Reverse())
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

        /// <summary>
        /// Sets the cell value as a shared string
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private int SetCellValue(string value)
        {
            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in ExcelFile.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == value)
                {
                    return i;
                }
                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            ExcelFile.WorkbookPart.SharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new Text(value)));
            ExcelFile.WorkbookPart.SharedStringTablePart.SharedStringTable.Save();
            return i;
        }

    }
}
