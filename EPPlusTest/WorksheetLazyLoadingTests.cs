using System;
using System.Drawing;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlusTest
{
    [TestClass]
    public class WorksheetLazyLoadingTests
    {
        public TestContext TestContext { get; set; }

        private readonly string[] testFruits = new[] { "Apple", "Peach", "Pear", "Banana" };
        private readonly DateTime testDateValue = new DateTime(2000, 1, 1, 3, 2, 1);

        private FileInfo CreateTestFile()
        {
            var file = Path.Combine(this.TestContext.TestResultsDirectory, this.TestContext.TestName + ".xlsx");

            if (File.Exists(file))
            {
                File.Delete(file);
            }

            return new FileInfo(file);
        }

        private ExcelWorksheet CreateTestSheet(ExcelWorkbook wb, int number)
        {
            var sheet = wb.Worksheets.Add("T" + number);

            sheet.Cells[1, 1].Value = this.testFruits[(number - 1) % this.testFruits.Length];
            sheet.Cells[1, 2].Value = number;
            sheet.Cells[1, 3].Value = number % 2 == 1;
            sheet.Cells[1, 4].Value = this.testDateValue.AddMinutes(number - 1);

            return sheet;
        }

        [TestMethod]
        public void ShouldNotLoadSheetsOnInitialization()
        {
            var testFile = this.CreateTestFile();

            using (var xl = new ExcelPackage(testFile))
            {
                this.CreateTestSheet(xl.Workbook, 1);
                this.CreateTestSheet(xl.Workbook, 2);
                this.CreateTestSheet(xl.Workbook, 3);

                xl.Save();
            }

            using (var xl = new ExcelPackage(testFile))
            {
                Assert.AreEqual(xl.Workbook.Worksheets.Count, 3);

                for (var i = 1; i <= xl.Workbook.Worksheets.Count; ++i)
                {
                    Assert.AreEqual(xl.Workbook.Worksheets[i].Name, "T" + i);
                    Assert.IsFalse(xl.Workbook.Worksheets[i].IsLoaded);
                }
            }
        }

        [TestMethod]
        public void ShouldNotTouchNotLoadedSheetsOnSave()
        {
            var testFile = this.CreateTestFile();

            using (var xl = new ExcelPackage(testFile))
            {
                this.CreateTestSheet(xl.Workbook, 1);

                xl.Save();
            }

            using (var xl = new ExcelPackage(testFile))
            {
                this.CreateTestSheet(xl.Workbook, 2);

                xl.Save();
            }

            using (var xl = new ExcelPackage(testFile))
            {
                var sheet1 = xl.Workbook.Worksheets["T1"];
                var sheet2 = xl.Workbook.Worksheets["T2"];

                Assert.AreEqual(sheet1.Cells[1, 1].Value, this.testFruits[0]);
                Assert.AreEqual(Convert.ToInt32(sheet1.Cells[1, 2].Value), 1);
                Assert.AreEqual(sheet1.Cells[1, 3].Value, true);
                Assert.AreEqual(DateTime.FromOADate((double)sheet1.Cells[1, 4].Value), this.testDateValue);

                Assert.AreEqual(sheet2.Cells[1, 1].Value, this.testFruits[1]);
                Assert.AreEqual(Convert.ToInt32(sheet2.Cells[1, 2].Value), 2);
                Assert.AreEqual(sheet2.Cells[1, 3].Value, false);
                Assert.AreEqual(DateTime.FromOADate((double)sheet2.Cells[1, 4].Value), this.testDateValue.AddMinutes(1));
            }
        }

        [TestMethod]
        public void ShouldNotTouchNotLoadedSheetsOnSaveWhenStylesApplied()
        {
            var testFile = this.CreateTestFile();

            using (var xl = new ExcelPackage(testFile))
            {
                var sheet = this.CreateTestSheet(xl.Workbook, 1);

                sheet.Column(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Column(1).Style.Fill.BackgroundColor.SetColor(Color.Blue);
                sheet.Cells[1, 2].Style.Numberformat.Format = "@";

                xl.Save();
            }

            using (var xl = new ExcelPackage(testFile))
            {
                var sheet = this.CreateTestSheet(xl.Workbook, 2);

                sheet.Column(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Column(1).Style.Fill.BackgroundColor.SetColor(Color.Red);
                sheet.Cells[1, 2].Style.Numberformat.Format = "0.00";

                xl.Save();
            }

            using (var xl = new ExcelPackage(testFile))
            {
                var sheet1 = xl.Workbook.Worksheets["T1"];
                var sheet2 = xl.Workbook.Worksheets["T2"];

                Assert.AreEqual(sheet1.Column(1).Style.Fill.PatternType, ExcelFillStyle.Solid);
                Assert.AreEqual(sheet1.Column(1).Style.Fill.BackgroundColor.Rgb, Color.Blue.ToArgb().ToString("X"));
                Assert.AreEqual(sheet1.Cells[1, 2].Style.Numberformat.Format, "@");

                Assert.AreEqual(sheet2.Column(1).Style.Fill.PatternType, ExcelFillStyle.Solid);
                Assert.AreEqual(sheet2.Column(1).Style.Fill.BackgroundColor.Rgb, Color.Red.ToArgb().ToString("X"));
                Assert.AreEqual(sheet2.Cells[1, 2].Style.Numberformat.Format, "0.00");
            }
        }

        [TestMethod]
        public void ShouldNWorkWithoutLazyLoadingWhenStylesApplied()
        {
            var testFile = this.CreateTestFile();

            using (var xl = new ExcelPackage(testFile))
            {
                {
                    var sheet = this.CreateTestSheet(xl.Workbook, 1);

                    sheet.Column(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Column(1).Style.Fill.BackgroundColor.SetColor(Color.Blue);
                    sheet.Cells[1, 2].Style.Numberformat.Format = "@";
                }

                {
                    var sheet = this.CreateTestSheet(xl.Workbook, 2);

                    sheet.Column(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Column(1).Style.Fill.BackgroundColor.SetColor(Color.Red);
                    sheet.Cells[1, 2].Style.Numberformat.Format = "0.00";
                }

                xl.Save();
            }

            using (var xl = new ExcelPackage(testFile))
            {
                var sheet1 = xl.Workbook.Worksheets["T1"];
                var sheet2 = xl.Workbook.Worksheets["T2"];

                Assert.AreEqual(sheet1.Column(1).Style.Fill.PatternType, ExcelFillStyle.Solid);
                Assert.AreEqual(sheet1.Column(1).Style.Fill.BackgroundColor.Rgb, Color.Blue.ToArgb().ToString("X"));
                Assert.AreEqual(sheet1.Cells[1, 2].Style.Numberformat.Format, "@");

                Assert.AreEqual(sheet2.Column(1).Style.Fill.PatternType, ExcelFillStyle.Solid);
                Assert.AreEqual(sheet2.Column(1).Style.Fill.BackgroundColor.Rgb, Color.Red.ToArgb().ToString("X"));
                Assert.AreEqual(sheet2.Cells[1, 2].Style.Numberformat.Format, "0.00");
            }
        }
    }
}