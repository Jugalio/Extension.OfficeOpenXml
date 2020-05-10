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
            file.Document.Close();
        }

    }
}
