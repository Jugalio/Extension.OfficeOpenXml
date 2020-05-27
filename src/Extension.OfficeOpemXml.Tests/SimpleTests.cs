using DocumentFormat.OpenXml.Spreadsheet;
using Extension.OfficeOpenXml.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Extension.OfficeOpemXml.Tests
{
    [TestFixture]
    public class SimpleTests: BaseTestClass
    {
        /// <summary>
        /// Simply creates an empty file with sheet1
        /// </summary>
        [Test]
        public void CreateEmptyFile()
        {
            var file = new ExcelFile();
            file.Create(GetGeneratedFilePath("EmptyFile.xlsx"));
            file.Save();
            file.Document.Close();
        }

        /// <summary>
        /// Simply creates an empty firl with sheet1
        /// </summary>
        [Test]
        public void CreateFileWithOneRow()
        {
            var file = new ExcelFile();
            file.Create(GetGeneratedFilePath("FileWithOneRow.xlsx"));
            var sheet = file.SheetList.FirstOrDefault();
            var row = sheet.AddRow();
            row.AddCell("Test 1");
            row.AddCell("Hallo");
            row.AddCell("5");
            row.AddCell(5);
            file.Save();

            var a = row.GetCellByColumnName("A").GetValue();

            file.Document.Close();
        }

        /// <summary>
        /// Loads a table and create copy
        /// </summary>
        [Test]
        public void LoadTable()
        {
            var file = new ExcelFile();
            var fileName = GetResourcesFilePath("Beispieltabelle.xlsx");
            file.Open(fileName, false);

            var copy = file.CopyWithStyle(GetGeneratedFilePath("Copied.xlsx"));
            copy.Save();
            copy.Document.Close();
        }

        /// <summary>
        /// Copy rows to a new file
        /// </summary>
        [Test]
        public void CopyRows()
        {
            var file = new ExcelFile();
            var fileName = GetResourcesFilePath("Beispieltabelle.xlsx");
            file.Open(fileName, false);
            var sourceSheet = file.SheetList.First();

            var copy = file.CopyWithStyle(GetGeneratedFilePath("CopiedRows.xlsx"), file.SheetList.First());
            var targetSheet = copy.SheetList.First();
            targetSheet.CopyRowFromOtherDocument(sourceSheet.Rows[0]);
            targetSheet.CopyRowFromOtherDocument(sourceSheet.Rows[2]);
            targetSheet.CopyRowFromOtherDocument(sourceSheet.Rows[3]);
            targetSheet.CopyRowFromOtherDocument(sourceSheet.Rows[5]);
            targetSheet.CopyRowFromOtherDocument(sourceSheet.Rows[6]);
            copy.Save();
            copy.Document.Close();
        }

    }
}
