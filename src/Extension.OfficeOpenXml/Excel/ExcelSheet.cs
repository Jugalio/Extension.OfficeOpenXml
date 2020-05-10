using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace Extension.OfficeOpenXml.Excel
{
    public class ExcelSheet
    {
        /// <summary>
        /// The excel file this sheet is a part of
        /// </summary>
        public ExcelFile ExcelFile;

        /// <summary>
        /// The open xml sheet instance
        /// </summary>
        public Sheet ThisSheet;

        /// <summary>
        /// The rows in this excel sheet
        /// </summary>
        public List<ExcelRow> Rows;

        /// <summary>
        /// Creates an empty sheet
        /// </summary>
        /// <param name="file"></param>
        /// <param name="name"></param>
        public void Create(ExcelFile file, string name, string id, uint sheetId)
        {
            ExcelFile = file;
            CreateEmptySheet(name, id, sheetId);
        }

        /// <summary>
        /// Adds a new row to the sheet
        /// </summary>
        /// <returns></returns>
        public ExcelRow AddRow()
        {
            var row = new ExcelRow();
            ExcelFile.SheetData.AppendChild(row.ThisRow);
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
