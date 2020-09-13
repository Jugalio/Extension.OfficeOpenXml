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
            var row = sheet.AppendNewRow();
            row.AppendCell("Test 1");
            row.AppendCell("Hallo");
            row.AppendCell("5");
            file.Save();

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
            targetSheet.AppendRow(sourceSheet.Rows[0], true);
            targetSheet.AppendRow(sourceSheet.Rows[1], true);
            targetSheet.AppendRow(sourceSheet.Rows[2], true);
            targetSheet.AppendRow(sourceSheet.Rows[3], true);
            targetSheet.AppendRow(sourceSheet.Rows[5], true);
            targetSheet.AppendRow(sourceSheet.Rows[6], true);
            targetSheet.AppendRow(sourceSheet.Rows[14], true);
            targetSheet.AppendRow(sourceSheet.Rows[15], true);
            copy.Save();
            copy.Document.Close();
        }

        [Test]
        public void GetCellsRightOf_Test()
        {
            var file = new ExcelFile();
            var fileName = GetResourcesFilePath("Beispieltabelle.xlsx");
            file.Open(fileName, false);
            var sourceSheet = file.SheetList.First();

            var cells = sourceSheet.Rows[1].GetCellsRightOf("D");
            Assert.IsTrue(cells.First().ThisCell.GetColumnName() == "E");
        }

        /// <summary>
        /// Copy rows to a new file
        /// </summary>
        [Test]
        public void GetColumn()
        {
            var file = new ExcelFile();
            var fileName = GetResourcesFilePath("Beispieltabelle.xlsx");
            file.Open(fileName, false);
            var sourceSheet = file.SheetList.First();
            sourceSheet.Rows.First().GetCellByColumnIndex(28);
        }

    }
}
