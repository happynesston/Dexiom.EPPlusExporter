﻿using Dexiom.EPPlusExporter;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dexiom.EPPlusExporterTests.Helpers;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using Xunit;
using Shouldly;

namespace EPPlusExporter.Tests
{
    public class TableExporterTests
    {
        [Fact]
        public void TableCreationTest()
        {
            const string tableName = "MyTable";
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var exporter = EnumerableExporter.Create(data);
            exporter.WorksheetName = tableName;
            exporter.TableStyle = TableStyles.Dark10;

            var excelPackage = exporter.CreateExcelPackage();

            var sheetToCheck = excelPackage.Workbook.Worksheets.Last();
            sheetToCheck.Tables[tableName].TableStyle.ShouldBe(exporter.TableStyle);
        }

        [Fact]
        public void AddWorksheetToNullPackageTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            ArgumentNullException ex = Assert.Throws<ArgumentNullException>(() => EnumerableExporter.Create(data).AppendToExcelPackage(null));
            ex.ParamName.ShouldBe("package");
        }

        [Fact]
        public void ConfigureTest()
        {
            var data = new[]
            {
                new { TextValue = "SomeText", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            var excelPackage = EnumerableExporter.Create(data)
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
            excelWorksheet.Cells[1, 1].Style.Fill.BackgroundColor.Rgb.ShouldNotBe("FFFF69B4");

            //data
            excelWorksheet.Cells[2, 1].Style.Fill.BackgroundColor.Rgb.ShouldBe("FFFF69B4");
        }

        [Fact]
        public void AutoFitColumnsTest()
        {
            var data = new[]
            {
                new { TextValue = "The quick brown fox jumps over the lazy dog", DateValue = DateTime.Now, DoubleValue = 10.2, IntValue = 5}
            };

            {
                var exporter = EnumerableExporter.Create(data, TableStyles.None)
                    .Configure(n => n.IntValue, configuration => configuration.Width = 60);
                exporter.AutoFitColumns = false;

                var excelPackage = exporter.CreateExcelPackage();
                //TestHelper.OpenDocument(excelPackage);

                var excelWorksheet = excelPackage.Workbook.Worksheets.First();

                //The two first columns should have the same size
                Math.Abs(excelWorksheet.Column(1).Width - excelWorksheet.Column(2).Width).ShouldBeLessThan(0.001);
                Math.Abs(excelWorksheet.Column(4).Width - 60).ShouldBeLessThan(0.001);
            }

            {
                var exporter = EnumerableExporter.Create(data, TableStyles.None)
                    .Configure(n => n.IntValue, configuration => configuration.Width = 60);
                exporter.AutoFitColumns = true;

                var excelPackage = exporter.CreateExcelPackage();
                //TestHelper.OpenDocument(excelPackage);

                var excelWorksheet = excelPackage.Workbook.Worksheets.First();

                var test1 = excelWorksheet.Column(1).Width;
                var test2 = excelWorksheet.Column(2).Width;
                var test4 = excelWorksheet.Column(4).Width;
                //The two first columns should have the same size
                Math.Abs(excelWorksheet.Column(1).Width - excelWorksheet.Column(2).Width).ShouldBeGreaterThan(0.001);
                Math.Abs(excelWorksheet.Column(4).Width - 60).ShouldBeLessThan(0.001);
            }
        }
    }
}