using Dexiom.EPPlusExporter;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporterTests.Extensions;
using Dexiom.EPPlusExporterTests.Helpers;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using Xunit;
using Shouldly;

namespace EPPlusExporter.Tests
{
    public class EnumerableExporterTests
    {
        [Fact]
        public void CreateExcelPackageTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelPackage = EnumerableExporter.Create(data).CreateExcelPackage();
            excelPackage.Workbook.Worksheets.Count.ShouldBe(1);
        }

        [Fact]
        public void AppendToExcelPackageTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelPackage = TestHelper.FakeAnExistingDocument();
            EnumerableExporter.Create(data)
                .CustomizeTable(range =>
                {
                    var newRange = range.Worksheet.Cells[range.End.Row, range.Start.Column, range.End.Row, range.End.Column];
                    newRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    newRange.Style.Fill.BackgroundColor.SetColor(Color.HotPink);
                })
                .AppendToExcelPackage(excelPackage);

            //TestHelper.OpenDocument(excelPackage);

            excelPackage.Workbook.Worksheets.Count.ShouldBe(2);
        }

        [Fact]
        public void ExportEmptyEnumerableTest()
        {
            var data = Enumerable.Empty<Tuple<string, int, bool>>();

            var excelPackage = EnumerableExporter.Create(data).CreateExcelPackage();
            //TestHelper.OpenDocument(excelPackage);

            excelPackage.ShouldNotBeNull();
        }

        [Fact]
        public void ExportNullTest()
        {
            IList<Tuple<string, int, bool>> data = null;

            // ReSharper disable once ExpressionIsAlwaysNull
            EnumerableExporter.Create(data).CreateExcelPackage().ShouldBeNull();
            // ReSharper disable once ExpressionIsAlwaysNull
            ObjectExporter.Create(data).AppendToExcelPackage(TestHelper.FakeAnExistingDocument()).ShouldBeNull();
        }
        
        [Fact]
        public void ConfigureTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var dynamicProperties = new[]
            {
                DynamicProperty.Create(data, "DynamicColumn1", "Display Name 1", typeof(DateTime?), n => DateTime.Now.AddDays(n.IntValue - 4)),
                DynamicProperty.Create(data, "DynamicColumn2", "Display Name 2", typeof(double), n => n.DoubleValue - 0.2)
            };

            var excelPackage = EnumerableExporter.Create(data, dynamicProperties)
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
                .Configure(new []{ "DynamicColumn1", "IntValue" }, n =>
                    {
                        n.Header.SetStyle = style =>
                        {
                            style.Font.Bold = true;
                            style.Font.Color.SetColor(Color.Black);
                        };
                    })
                .CustomizeTable(range =>
                {
                    var newRange = range.Worksheet.Cells[range.End.Row, range.Start.Column, range.End.Row, range.End.Column];
                    newRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    newRange.Style.Fill.BackgroundColor.SetColor(Color.HotPink);
                })
                .CreateExcelPackage();

            //TestHelper.OpenDocument(excelPackage);


            var excelWorksheet = excelPackage.Workbook.Worksheets.First();

            //header
            excelWorksheet.Cells[1, 2].Style.Border.Bottom.Style.ShouldBe(ExcelBorderStyle.Thick);
            excelWorksheet.Cells[1, 2].Style.Border.Bottom.Style.ShouldBe(ExcelBorderStyle.Thick);
            excelWorksheet.Cells[1, 2].Text.ShouldBe(" ");
            excelWorksheet.Cells[1, 4].Text.ShouldBe("Int Value");
            excelWorksheet.Cells[1, 1].Style.Fill.BackgroundColor.Rgb.ShouldNotBe("FFFF69B4");

            //data
            excelWorksheet.Cells[2, 2].Text.ShouldBe(DateTime.Now.ToString("dd-MM-yyyy"));
            excelWorksheet.Cells[2, 2].Style.Border.Left.Style.ShouldBe(ExcelBorderStyle.Dashed);
            excelWorksheet.Cells[2, 2].Style.Border.Right.Style.ShouldBe(ExcelBorderStyle.Dashed);
            excelWorksheet.Cells[2, 1].Style.Fill.BackgroundColor.Rgb.ShouldBe("FFFF69B4");
        }

        #region Fluent Interface Tests

        //[Fact]
        //public void WorksheetConfigurationTest()
        //{
        //    const string newWorksheetName = "1 - NewSheet";
        //    const string newWorksheetExpectedTableName = "t1_-_NewSheet";
        //    var data = new[]
        //    {
        //        new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
        //    };

        //    var test = IsValidName(newWorksheetExpectedTableName);

        //    var excelPackage = TestHelper.FakeAnExistingDocument();
        //    var eporter = EnumerableExporter.Create(data);

        //    //set properties
        //    eporter.WorksheetName = newWorksheetName;
        //    var sheetToCheck = eporter.AppendToExcelPackage(excelPackage);

        //    //TestHelper.OpenDocument(excelPackage);

        //    //check properties
        //    sheetToCheck.Name.ShouldBe(newWorksheetName);
        //    sheetToCheck.Tables[newWorksheetExpectedTableName].ShouldNotBeNull();
            
        //}

        [Fact]
        public void DefaultNumberFormatTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var dynamicProperties = new[]
            {
                DynamicProperty.Create(data, "DynamicColumn1", "Display Name 1", typeof(DateTime?), n => DateTime.Now.AddDays(n.IntValue - 4)),
                DynamicProperty.Create(data, "DynamicColumn2", "Display Name 2", typeof(double), n => n.DoubleValue - 0.2)
            };


            var excelPackage = EnumerableExporter.Create(data, dynamicProperties)
                .DefaultNumberFormat(typeof(DateTime), "| yyyy-MM-dd")
                .DefaultNumberFormat(typeof(DateTime?), "|| yyyy-MM-dd")
                .DefaultNumberFormat(typeof(double), "0.00 $")
                .DefaultNumberFormat(typeof(int), "00")
                .CreateExcelPackage();
            var excelWorksheet = excelPackage.Workbook.Worksheets.First();

            //TestHelper.OpenDocument(excelPackage);

            string numberDecimalSeparator = NumberFormatInfo.CurrentInfo.NumberDecimalSeparator;

            excelWorksheet.Cells[2, 2].Text.ShouldBe(DateTime.Today.ToString("| yyyy-MM-dd")); //DateValue
            excelWorksheet.Cells[2, 3].Text.ShouldBe($"10{numberDecimalSeparator}20 $"); //DoubleValue
            excelWorksheet.Cells[2, 4].Text.ShouldBe("05"); //IntValue
            excelWorksheet.Cells[2, 5].Text.ShouldBe(DateTime.Today.AddDays(1).ToString("|| yyyy-MM-dd")); //DynamicColumn1
            excelWorksheet.Cells[2, 6].Text.ShouldBe($"10{numberDecimalSeparator}00 $"); //DynamicColumn2

        }

        [Fact]
        public void NumberFormatForTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelWorksheet = EnumerableExporter.Create(data)
                .NumberFormatFor(n => n.DateValue, "DATE: yyyy-MM-dd")
                .NumberFormatFor(n => n.DoubleValue, "0.00 $")
                .NumberFormatFor(n => n.IntValue, "00")
                .CreateExcelPackage()
                .Workbook.Worksheets.First();

            string numberDecimalSeparator = NumberFormatInfo.CurrentInfo.NumberDecimalSeparator;

            excelWorksheet.Cells[2, 2].Text.ShouldBe(DateTime.Today.ToString("DATE: yyyy-MM-dd")); //DateValue
            excelWorksheet.Cells[2, 3].Text.ShouldBe($"10{numberDecimalSeparator}20 $"); //DoubleValue
            excelWorksheet.Cells[2, 4].Text.ShouldBe("05"); //IntValue
        }

        [Fact]
        public void DisplayTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelWorksheet = EnumerableExporter.Create(data)
                .Ignore(n => n.DateValue)
                .Display(n => new
                {
                    n.TextValue,
                    n.DoubleValue
                })
                .CreateExcelPackage()
                .Workbook.Worksheets.First();

            excelWorksheet.Cells[1, 1].Text.ShouldBe("Text Value");
            excelWorksheet.Cells[1, 2].Text.ShouldBe("Double Value");
            excelWorksheet.Cells[1, 3].Text.ShouldBe(string.Empty);
        }

        [Fact]
        public void IgnoreTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelWorksheet = EnumerableExporter.Create(data)
                .Ignore(n => n.TextValue)
                .CreateExcelPackage()
                .Workbook.Worksheets.First();

            excelWorksheet.Cells[1, 1].Text.ShouldBe("Date Value");
        }


        [Fact]
        public void TextFormatForTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };


            const string textFormat = "Prefix: {0}";
            const string dateFormat = "{0:yyyy-MM-dd HH:mm}";
            var exporter = EnumerableExporter.Create(data)
                .TextFormatFor(n => n.TextValue, textFormat)
                .TextFormatFor(n => n.DateValue, dateFormat);

            var excelPackage = exporter.CreateExcelPackage();
            var excelWorksheet = excelPackage.Workbook.Worksheets.First();

            excelWorksheet.Cells[2, 1].Text.ShouldBe(string.Format(textFormat, data.First().TextValue)); //TextValue
            excelWorksheet.Cells[2, 2].Text.ShouldBe(string.Format(dateFormat, data.First().DateValue)); //DateValue
            excelWorksheet.Cells[2, 2].Value.ToString().ShouldBe(string.Format(dateFormat, data.First().DateValue)); //DateValue
        }

        [Fact]
        public void StyleForTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };
            
            const string dateFormat = "yyyy-MM-dd HH:mm";
            var exporter = EnumerableExporter.Create(data)
                .StyleFor(n => n.DateValue, n => n.Numberformat.Format = dateFormat);

            var excelPackage = exporter.CreateExcelPackage();
            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            excelWorksheet.Cells[2, 2].Text.ShouldBe(data.First().DateValue.ToString(dateFormat)); //DateValue
        }

        [Fact]
        public void HeaderStyleForTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };
    
            var exporter = EnumerableExporter.Create(data)
                .HeaderStyleFor(n => new { n.DateValue, n.DoubleValue, n.IntValue }, 
                    style => style.Border.Bottom.Style = ExcelBorderStyle.Thick);

            var excelPackage = exporter.CreateExcelPackage();
            //TestHelper.OpenDocument(excelPackage);

            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            excelWorksheet.Cells[1, 2].Style.Border.Bottom.Style.ShouldBe(ExcelBorderStyle.Thick);
        }
        
        [Fact]
        public void ConditionalStyleForTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText0", DateValue = DateTime.Now, DoubleValue = 0, IntValue = 5},
                new { TextValue = "SomeText1", DateValue = DateTime.Now, DoubleValue = 1, IntValue = 5},
                new { TextValue = "SomeText2", DateValue = DateTime.Now, DoubleValue = 2, IntValue = 5},
                new { TextValue = "SomeText3", DateValue = DateTime.Now, DoubleValue = 3, IntValue = 5}
            };

            var exporter = EnumerableExporter.Create(data)
                .ConditionalStyleFor(n => n.DoubleValue, (entry, style) =>
                {
                    if (entry.DoubleValue > 1)
                    {
                        style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                    }
                });

            var excelPackage = exporter.CreateExcelPackage();
            //TestHelper.OpenDocument(excelPackage);

            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            excelWorksheet.Cells[3, 3].Style.Border.Bottom.Style.ShouldBe(ExcelBorderStyle.None);
            excelWorksheet.Cells[4, 3].Style.Border.Bottom.Style.ShouldBe(ExcelBorderStyle.Thick);
        }

        [Fact]
        public void FormulaForTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText0", DateValue = DateTime.Now, DoubleValue = 0, IntValue = 5},
                new { TextValue = "SomeText1", DateValue = DateTime.Now, DoubleValue = 1, IntValue = 5},
                new { TextValue = "SomeText2", DateValue = DateTime.Now, DoubleValue = 2, IntValue = 5},
                new { TextValue = "SomeText3", DateValue = DateTime.Now, DoubleValue = 3, IntValue = 5}
            };

            var exporter = EnumerableExporter.Create(data)
                .FormulaFor(n => n.TextValue, (row, value) => $"=\"Text=\" & \"{value}-\" & \"{row.DoubleValue:0.00}\"");

            var excelPackage = exporter.CreateExcelPackage();
            //TestHelper.OpenDocument(excelPackage);

            var excelWorksheet = excelPackage.Workbook.Worksheets.First();
            excelWorksheet.Cells[3, 1].Formula.ShouldBe("=\"Text=\" & \"SomeText1-\" & \"1,00\"");
            excelWorksheet.Cells[4, 1].Formula.ShouldBe("=\"Text=\" & \"SomeText2-\" & \"2,00\"");
        }

        #endregion
    }
}