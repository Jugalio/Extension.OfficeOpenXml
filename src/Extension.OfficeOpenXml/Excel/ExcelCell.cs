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
        /// The name of the column in whích the cell will be added
        /// </summary>
        public string ColumnName;

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
        /// Creates a new excel from a given cell
        /// </summary>
        /// <param name="value"></param>
        public ExcelCell(ExcelFile file, ExcelCell refCell)
        {
            ExcelFile = file;
            ColumnName = refCell.ThisCell.GetColumnName();

            if (refCell.ThisCell.CellFormula != null)
            {
                //If a cell formula is defined we do not set the cell value
                ThisCell = new Cell();
            }
            else if (refCell.ThisCell.DataType == null)
            {
                ThisCell = new Cell()
                {
                    CellValue = new CellValue(refCell.Value),
                };
            }
            else if (refCell.ThisCell.DataType == CellValues.SharedString)
            {
                var index = SetSharedCellValue(refCell.Value);
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
        public ExcelCell(ExcelFile file, Cell sdkCell)
        {
            ExcelFile = file;
            ColumnName = sdkCell.GetColumnName();

            ThisCell = sdkCell;
        }

        /// <summary>
        /// Creates a new excel cell with the given value
        /// </summary>
        /// <param name="value"></param>
        public ExcelCell(ExcelFile file, string columnName, string value)
        {
            ExcelFile = file;
            ColumnName = columnName;

            ThisCell = new Cell()
            {
                CellValue = new CellValue(value),
            };
        }

        /// <summary>
        /// Moves a cell one column to the right
        /// After this the SetCellRef has to be called
        /// </summary>
        internal void MoveRight(uint rowIndex)
        {
            var oldCellRef = ThisCell.CellReference;
            ColumnName = ColumnName.IterateUpperLetter();
            SetCellRef(rowIndex);

            //If the cell includes a formula, we also have to update the calculation cell
            if (ThisCell.CellFormula != null)
            {
                ThisCell.CellValue = null;
                var calculationChainPart = ExcelFile.WorkbookPart.CalculationChainPart;
                var calculationChain = calculationChainPart.CalculationChain;
                var calculationCells = calculationChain.Elements<CalculationCell>().ToList();
                var calculationCell = calculationCells.Where(calcCell => calcCell.CellReference == oldCellRef).FirstOrDefault();
                if (calculationCell != null)
                {
                    calculationCell.CellReference = ThisCell.CellReference;
                }
            }
        }

        /// <summary>
        /// Moves a cell one column to the right
        /// After this the SetCellRef has to be called
        /// </summary>
        internal void MoveLeft(uint rowIndex)
        {
            var oldCellRef = ThisCell.CellReference;
            ColumnName = ColumnName.ReverseIterateUpperLetter();
            SetCellRef(rowIndex);

            //If the cell includes a formula, we also have to update the calculation cell
            if (ThisCell.CellFormula != null)
            {
                ThisCell.CellValue = null;
                var calculationChainPart = ExcelFile.WorkbookPart.CalculationChainPart;
                var calculationChain = calculationChainPart.CalculationChain;
                var calculationCells = calculationChain.Elements<CalculationCell>().ToList();
                var calculationCell = calculationCells.Where(calcCell => calcCell.CellReference == oldCellRef).FirstOrDefault();
                if (calculationCell != null)
                {
                    calculationCell.CellReference = ThisCell.CellReference;
                }
            }
        }

        /// <summary>
        /// Sets the cell reference with the row to which this cell is added
        /// </summary>
        /// <param name="rowIndex"></param>
        public void SetCellRef(uint rowIndex)
        {
            ThisCell.CellReference = $"{ColumnName}{rowIndex}";
        }

        /// <summary>
        /// Sets a string value
        /// </summary>
        /// <param name="value"></param>
        public void SetValue(string value)
        {
            var index = SetSharedCellValue(value);
            ThisCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            ThisCell.CellValue = new CellValue(index.ToString());
        }

        /// <summary>
        /// Sets a int value
        /// </summary>
        /// <param name="value"></param>
        public void SetValue(int value)
        {
            ThisCell.CellValue = new CellValue(value.ToString());
        }

        /// <summary>
        /// Sets a double value
        /// </summary>
        /// <param name="value"></param>
        public void SetValue(double value)
        {
            ThisCell.CellValue = new CellValue(value.ToString());
        }

        /// <summary>
        /// Sets a long value
        /// </summary>
        /// <param name="value"></param>
        public void SetValue(long value)
        {
            ThisCell.CellValue = new CellValue(value.ToString());
        }

        /// <summary>
        /// Sets a formula for this cell
        /// </summary>
        /// <param name="value"></param>
        public void SetFormual(string value)
        {
            ThisCell.CellFormula = new CellFormula(value);
            ThisCell.DataType = null;
        }

        /// <summary>
        /// Sets the cell value as a shared string
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private int SetSharedCellValue(string value)
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
