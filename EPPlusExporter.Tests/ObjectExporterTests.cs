using Dexiom.EPPlusExporter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporterTests.Helpers;
using OfficeOpenXml.Style;
using System.Globalization;
using Xunit;
using Shouldly;

namespace EPPlusExporter.Tests
{
    public class ObjectExporterTests
    {
        [Fact]
        public void CreateExcelPackageTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var excelPackage = ObjectExporter.Create(data).CreateExcelPackage();
            excelPackage.Workbook.Worksheets.Count.ShouldBe(1);
        }

        [Fact]
        public void AppendToExcelPackageTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var excelPackage = TestHelper.FakeAnExistingDocument();
            ObjectExporter.Create(data).AppendToExcelPackage(excelPackage);
            excelPackage.Workbook.Worksheets.Count.ShouldBe(2);
            //TestHelper.OpenDocumentIfRequired(excelPackage);
        }

        [Fact]
        public void ExportNullTest()
        {
            IList<Tuple<string, int, bool>> data = null;

            // ReSharper disable once ExpressionIsAlwaysNull
            ObjectExporter.Create(data).CreateExcelPackage().ShouldBeNull();
            // ReSharper disable once ExpressionIsAlwaysNull
            ObjectExporter.Create(data).AppendToExcelPackage(TestHelper.FakeAnExistingDocument()).ShouldBeNull();
        }

        [Fact]
        public void ConfigureTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };
            
            var excelPackage = ObjectExporter.Create(data)
                .Configure(n => n.IntValue, configuration =>
                {
                    configuration.Header.Text = "";
                })
                .Configure(n => n.DateValue, configuration =>
                {
                    configuration.Header.Text = " ";
                    configuration.Header.SetStyle = style =>
                    {
                        style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                    };
                    configuration.Content.NumberFormat = "dd-MM-yyyy";
                    configuration.Content.SetStyle = style =>
                    {
                        style.Border.Left.Style = ExcelBorderStyle.Dashed;
                        style.Border.Right.Style = ExcelBorderStyle.Dashed;
                    };
                })
                .CreateExcelPackage();

            //TestHelper.OpenDocument(excelPackage);

            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            //header
            excelWorksheet.Cells[3, 1].Style.Border.Bottom.Style.ShouldBe(ExcelBorderStyle.Thick);
            excelWorksheet.Cells[3, 1].Text.ShouldBe(" ");
            excelWorksheet.Cells[5, 1].Text.ShouldBe("Int Value");

            //data
            excelWorksheet.Cells[3, 2].Text.ShouldBe(DateTime.Now.ToString("dd-MM-yyyy"));
            excelWorksheet.Cells[3, 2].Style.Border.Left.Style.ShouldBe(ExcelBorderStyle.Dashed);
            excelWorksheet.Cells[3, 2].Style.Border.Right.Style.ShouldBe(ExcelBorderStyle.Dashed);
        }

        [Fact]
        public void WorksheetConfigurationTest()
        {
            const string newWorksheetName = "NewSheet";
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelPackage = TestHelper.FakeAnExistingDocument();
            var eporter = ObjectExporter.Create(data);

            //set properties
            eporter.WorksheetName = newWorksheetName;
            eporter.AppendToExcelPackage(excelPackage);

            //check properties
            var sheetToCheck = excelPackage.Workbook.Worksheets.Last();
            sheetToCheck.Name.ShouldBe(newWorksheetName);
        }

        [Fact]
        public void DefaultNumberFormatTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var excelWorksheet = ObjectExporter.Create(data)
                .DefaultNumberFormat(typeof(DateTime), "DATE: yyyy-MM-dd")
                .DefaultNumberFormat(typeof(double), "0.00 $")
                .DefaultNumberFormat(typeof(int), "00")
                .CreateExcelPackage()
                .Workbook.Worksheets.First();

            string numberDecimalSeparator = NumberFormatInfo.CurrentInfo.NumberDecimalSeparator;

            excelWorksheet.Cells[3, 2].Text.ShouldBe(DateTime.Today.ToString("DATE: yyyy-MM-dd")); //DateValue
            excelWorksheet.Cells[4, 2].Text.ShouldBe($"10{numberDecimalSeparator}20 $"); //DoubleValue
            excelWorksheet.Cells[5, 2].Text.ShouldBe("05"); //IntValue
        }

        [Fact]
        public void NumberFormatForTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var excelWorksheet = ObjectExporter.Create(data)
                .NumberFormatFor(n => n.DateValue, "DATE: yyyy-MM-dd")
                .NumberFormatFor(n => n.DoubleValue, "0.00 $")
                .NumberFormatFor(n => n.IntValue, "00")
                .CreateExcelPackage()
                .Workbook.Worksheets.First();

            string numberDecimalSeparator = NumberFormatInfo.CurrentInfo.NumberDecimalSeparator;

            excelWorksheet.Cells[3, 2].Text.ShouldBe(DateTime.Today.ToString("DATE: yyyy-MM-dd")); //DateValue
            excelWorksheet.Cells[4, 2].Text.ShouldBe($"10{numberDecimalSeparator}20 $"); //DoubleValue
            excelWorksheet.Cells[5, 2].Text.ShouldBe("05"); //IntValue
        }

        [Fact]
        public void DisplayTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var excelWorksheet = ObjectExporter.Create(data)
                .Ignore(n => n.DateValue)
                .Display(n => new
                {
                    n.TextValue,
                    n.DoubleValue
                })
                .CreateExcelPackage()
                .Workbook.Worksheets.First();

            excelWorksheet.Cells[2, 1].Text.ShouldBe("Text Value");
            excelWorksheet.Cells[3, 1].Text.ShouldBe("Double Value");
            excelWorksheet.Cells[4, 1].Text.ShouldBe(string.Empty);
        }
        
        [Fact]
        public void IgnoreTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var excelWorksheet = ObjectExporter.Create(data)
                .Ignore(n => n.TextValue)
                .CreateExcelPackage()
                .Workbook.Worksheets.First();

            excelWorksheet.Cells[2, 1].Text.ShouldBe("Date Value");
        }
        
        [Fact]
        public void TextFormatForTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            const string textFormat = "Prefix: {0}";
            const string dateFormat = "{0:yyyy-MM-dd HH:mm}";
            var exporter = ObjectExporter.Create(data)
                .TextFormatFor(n => n.TextValue, textFormat)
                .TextFormatFor(n => n.DateValue, dateFormat);

            var excelPackage = exporter.CreateExcelPackage();
            var excelWorksheet = excelPackage.Workbook.Worksheets.First();

            excelWorksheet.Cells[2, 2].Text.ShouldBe(string.Format(textFormat, data.TextValue)); //TextValue
            excelWorksheet.Cells[3, 2].Text.ShouldBe(string.Format(dateFormat, data.DateValue)); //DateValue
            excelWorksheet.Cells[3, 2].Value.ToString().ShouldBe(string.Format(dateFormat, data.DateValue)); //DateValue
        }

        [Fact]
        public void StyleForTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };
            
            const string dateFormat = "yyyy-MM-dd HH:mm";
            var exporter = ObjectExporter.Create(data)
                .StyleFor(n => n.DateValue, n => n.Numberformat.Format = dateFormat);

            var excelPackage = exporter.CreateExcelPackage();
            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            
            excelWorksheet.Cells[3, 2].Text.ShouldBe(data.DateValue.ToString(dateFormat)); //DateValue
        }

        [Fact]
        public void HeaderStyleForTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var exporter = ObjectExporter.Create(data)
                .HeaderStyleFor(n => new { n.DateValue, n.DoubleValue, n.IntValue },
                style => style.Border.Right.Style = ExcelBorderStyle.Thick);

            var excelPackage = exporter.CreateExcelPackage();
            //TestHelper.OpenDocument(excelPackage);

            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            excelWorksheet.Cells[3, 1].Style.Border.Right.Style.ShouldBe(ExcelBorderStyle.Thick);
            excelWorksheet.Cells[4, 1].Style.Border.Right.Style.ShouldBe(ExcelBorderStyle.Thick);
            excelWorksheet.Cells[5, 1].Style.Border.Right.Style.ShouldBe(ExcelBorderStyle.Thick);
        }

        [Fact]
        public void ConditionalStyleForTest()
        {
            var data = new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5 };

            var exporter = ObjectExporter.Create(data)
                .ConditionalStyleFor(n => n.DateValue, (entry, style) =>
                {
                    if (entry.DoubleValue < 1)
                    {
                        style.Border.Bottom.Style = ExcelBorderStyle.Dashed;
                    }
                })
                .ConditionalStyleFor(n => n.DoubleValue, (entry, style) =>
                {
                    if (entry.DoubleValue > 1)
                    {
                        style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                    }
                })
                .ConditionalStyleFor(n => n.IntValue, (entry, style) =>
                {
                    if (entry.DoubleValue > 1)
                    {
                        style.Border.Bottom.Style = ExcelBorderStyle.Dashed;
                    }
                });

            var excelPackage = exporter.CreateExcelPackage();
            //TestHelper.OpenDocument(excelPackage);

            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            excelWorksheet.Cells[3, 1].Style.Border.Bottom.Style.ShouldBe(ExcelBorderStyle.None);
            excelWorksheet.Cells[4, 1].Style.Border.Bottom.Style.ShouldBe(ExcelBorderStyle.Thick);
            excelWorksheet.Cells[5, 2].Style.Border.Bottom.Style.ShouldBe(ExcelBorderStyle.Dashed);
        }
    }
}