using DocumentFormat.OpenXml.Spreadsheet;
using Extension.Utilities.ClassExtensions;
using System;
using System.Collections.Generic;
using System.Globalization;
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
                Cells.Add(new ExcelCell(ExcelFile, cell));
            }
        }

        /// <summary>
        /// Adds a new cell to the row
        /// </summary>
        /// <param name="value"></param>
        public ExcelCell AppendCell(string value)
        {
            var column = Cells.LastOrDefault()?.ThisCell.GetNextColumnName() ?? "A";
            var cell = new ExcelCell(ExcelFile, column, value);
            AppendCell(cell);
            return cell;
        }

        /// <summary>
        /// Adds a new cell to the row
        /// </summary>
        /// <param name="value"></param>
        public ExcelCell AppendCell(ExcelCell cell, bool clone = false)
        {
            if (clone)
            {
                cell = new ExcelCell(ExcelFile, cell);
            }
            var column = Cells.LastOrDefault()?.ThisCell.GetNextColumnName() ?? "A";
            cell.ColumnName = column;
            cell.SetCellRef(RowIndex);
            ThisRow.Append(cell.ThisCell);
            Cells.Add(cell);
            return cell;
        }

        /// <summary>
        /// Removes a given cell from the row
        /// </summary>
        /// <param name="cell"></param>
        public void RemoveCell(ExcelCell cell, bool moveAllRightOf = false)
        {
            if (cell != null)
            {
                if (Cells.Contains(cell))
                {
                    Cells.Remove(cell);
                    cell.ThisCell.Remove();
                }
                if (moveAllRightOf)
                {
                    foreach (var c in GetCellsRightOf(cell.ColumnName))
                    {
                        c.MoveLeft(RowIndex);
                    }
                }
            }
        }

        /// <summary>
        /// Removes a given cell from the row
        /// </summary>
        /// <param name="cell"></param>
        public void RemoveCell(string columnId, bool moveAllRightOf = false)
        {
            RemoveCell(GetCellByColumnName(columnId), false);
            if (moveAllRightOf)
            {
                foreach (var c in GetCellsRightOf(columnId))
                {
                    c.MoveLeft(RowIndex);
                }
            }
        }

        /// <summary>
        /// Removes a given cell from the row
        /// </summary>
        /// <param name="cell"></param>
        public void RemoveCell(uint columnIndex, bool moveAllRightOf = false)
        {
            RemoveCell(GetCellByColumnIndex(columnIndex), false);
            if (moveAllRightOf)
            {
                foreach (var c in GetCellsRightOf(GetColumnId(columnIndex)))
                {
                    c.MoveLeft(RowIndex);
                }
            }
        }

        /// <summary>
        /// Adds a new cell to the row without changing the column name
        /// </summary>
        /// <param name="value"></param>
        public ExcelCell InsertCell(ExcelCell cell, bool clone = false)
        {
            return InsertCellAt(cell, cell.ColumnName, clone);
        }

        /// <summary>
        /// Adds a new cell to the row without changing the column name
        /// </summary>
        /// <param name="value"></param>
        public ExcelCell InsertCellAt(ExcelCell cell, uint columnIndex, bool clone = false)
        {
            return InsertCellAt(cell, GetColumnId(columnIndex), clone);
        }

        /// <summary>
        /// Adds a new cell to the row without changing the column name
        /// </summary>
        /// <param name="value"></param>
        public ExcelCell InsertCellAt(ExcelCell cell, string columnId, bool clone = false)
        {
            if (clone)
            {
                cell = new ExcelCell(ExcelFile, cell);
            }

            //If there already is a cell defined for that place move all cells right of it one to the right
            var occupied = GetCellByColumnName(columnId);
            var toTheRight = GetCellsRightOf(columnId).ToList();
            var firstToTheRight = toTheRight.FirstOrDefault();
            if (occupied != null)
            {
                occupied.MoveRight(RowIndex);
                toTheRight.ForEach(c =>
                {
                    c.MoveRight(RowIndex);
                });
            }

            cell.ColumnName = columnId;
            cell.SetCellRef(RowIndex);
            if (firstToTheRight != null)
            {
                ThisRow.InsertBefore(cell.ThisCell, occupied.ThisCell);
            }
            else
            {
                ThisRow.Append(cell.ThisCell);
            }
            Cells.Add(cell);
            return cell;
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
            var id = GetColumnId(index);
            return Cells.FirstOrDefault(c => c.ThisCell.GetColumnName() == id);
        }

        /// <summary>
        /// Append a list of cells to the end
        /// </summary>
        /// <param name="cells"></param>
        public void AppendCells(List<ExcelCell> cells, bool clone = false)
        {
            cells.ForEach(c => AppendCell(c, clone));
        }

        /// <summary>
        ///Insert Cells, without changing their columnname
        /// </summary>
        /// <param name="cells"></param>
        public void InsertCells(List<ExcelCell> cells, bool clone = false)
        {
            cells.ForEach(c => InsertCell(c, clone));
        }

        /// <summary>
        /// Get all cells right of the provided columnIndex
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public IEnumerable<ExcelCell> GetCellsRightOf(uint columnIndex)
        {
            return GetCellsRightOf(GetColumnId(columnIndex));
        }

        /// <summary>
        /// Get all cells right of the provided columnname
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public IEnumerable<ExcelCell> GetCellsRightOf(string columnId)
        {
            return Cells.Where(c => c.ThisCell.GetColumnName().CompareTo(columnId) > 0);
        }

        /// <summary>
        /// Moves a full column within this row
        /// </summary>
        /// <param name="currentColumnIndex"></param>
        /// <param name="newColumnIndex"></param>
        internal void MoveColumn(uint currentColumnIndex, uint newColumnIndex)
        {
            MoveColumn(GetColumnId(currentColumnIndex), GetColumnId(newColumnIndex));
        }

        /// <summary>
        /// Moves a full column within this row
        /// </summary>
        /// <param name="currentColumnId"></param>
        /// <param name="newColumnId"></param>
        internal void MoveColumn(string currentColumnId, string newColumnId)
        {
            if (currentColumnId != newColumnId)
            {
                var refCell = GetCellByColumnName(currentColumnId);
                if (refCell != null)
                {
                    RemoveCell(refCell, true);
                    InsertCellAt(refCell, newColumnId, true);
                }
            }
        }

        /// <summary>
        /// Get the columnid for the index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        private string GetColumnId(uint index)
        {
            string id = "A";

            for (uint i = 1; i < index; i++)
            {
                id = id.IterateUpperLetter();
            }

            return id;
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
