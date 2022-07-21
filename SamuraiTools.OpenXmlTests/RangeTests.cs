using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SamuraiTools.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SamuraiTools.OpenXml.Spreadsheet.Tests
{
    [TestClass()]
    public class RangeTests
    {
        [TestMethod()]
        [DataRow(true)]
        [DataRow(false)]
        public void MergeCellsTest(bool center)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1", DataType = CellValues.Number, CellValue = new CellValue(0) }, new Cell() { CellReference = "B1", DataType = CellValues.Number, CellValue = new CellValue(0) }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                var range = new Range(worksheet, 1, 1, 1, 2);

                range.MergeCells(center);

                Assert.AreEqual("A1:B1", worksheet.GetFirstChild<MergeCells>().GetFirstChild<MergeCell>().Reference.Value);
                Assert.IsFalse(worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().Elements<Cell>().Last().HasChildren);
                if (center)
                {
                    var cellFormat = document.WorkbookPart.WorkbookStylesPart.Stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>().StyleIndex.Value);
                    Assert.AreEqual(HorizontalAlignmentValues.Center, cellFormat.Alignment.Horizontal.Value);
                }
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void AutoFitTest()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1", DataType = CellValues.Number, CellValue = new CellValue(0) }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font() { FontName = new FontName() { Val = "Calibri" }, FontSize = new FontSize() { Val = 11D } }), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                var range = new Range(worksheet, null, 1);

                range.AutoFit();

                var column = worksheet.GetFirstChild<Columns>().GetFirstChild<Column>();
                Assert.IsTrue(column.CustomWidth);
                Assert.IsTrue(column.BestFit);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void ApplyCellStyleTestCellRange()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }, new Cell() { CellReference = "B1" }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                var range = new Range(worksheet, 1, 1, 1, 2);

                range.ApplyCellStyle(BuiltinCellStyle.Hyperlink);

                var cell1 = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cell2 = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().Elements<Cell>().Last();
                var cellStyle = document.WorkbookPart.WorkbookStylesPart.Stylesheet.GetFirstChild<CellStyles>().Elements<CellStyle>().Single(s => s.BuiltinId == (uint)BuiltinCellStyle.Hyperlink);
                var cellFormat = document.WorkbookPart.WorkbookStylesPart.Stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell1.StyleIndex.Value);
                
                Assert.AreEqual(cellStyle.FormatId.Value, cellFormat.FormatId.Value);
                Assert.AreEqual(cell1.StyleIndex.Value, cell2.StyleIndex.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void ApplyCellStyleTestRowRange()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }, new Cell() { CellReference = "B1" }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                var range = new Range(worksheet, 1, null);

                range.ApplyCellStyle(BuiltinCellStyle.Hyperlink);

                var row = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>();
                var cellStyle = document.WorkbookPart.WorkbookStylesPart.Stylesheet.GetFirstChild<CellStyles>().Elements<CellStyle>().Single(s => s.BuiltinId == (uint)BuiltinCellStyle.Hyperlink);
                var cellFormat = document.WorkbookPart.WorkbookStylesPart.Stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)row.StyleIndex.Value);

                Assert.AreEqual(cellStyle.FormatId.Value, cellFormat.FormatId.Value);
                Assert.IsTrue(row.Elements<Cell>().All(c => c.StyleIndex.Value == row.StyleIndex.Value));
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void ApplyCellStyleTestColumnRange()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }, new Cell() { CellReference = "B1" }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                var range = new Range(worksheet, null, 1);

                range.ApplyCellStyle(BuiltinCellStyle.Hyperlink);

                var column = worksheet.GetFirstChild<Columns>().GetFirstChild<Column>();
                var cell1 = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cell2 = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().Elements<Cell>().Last();
                var cellStyle = document.WorkbookPart.WorkbookStylesPart.Stylesheet.GetFirstChild<CellStyles>().Elements<CellStyle>().Single(s => s.BuiltinId == (uint)BuiltinCellStyle.Hyperlink);
                var cellFormat = document.WorkbookPart.WorkbookStylesPart.Stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)column.Style.Value);

                Assert.AreEqual(cellStyle.FormatId.Value, cellFormat.FormatId.Value);
                Assert.AreEqual(column.Style.Value, cell1.StyleIndex.Value);
                Assert.AreNotEqual(column.Style.Value, cell2.StyleIndex?.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow("Helvetica")]
        [DataRow(null)]
        public void SetFontTestName(string name)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font() { FontName = new FontName() { Val = "Calibri" } }), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                var range = new Range(worksheet, 1, 1);

                range.SetFont(name: name);

                var cell = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cellFormat = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell.StyleIndex.Value);
                var font = stylesheet.GetFirstChild<Fonts>().Elements<Font>().ElementAt((int)cellFormat.FontId.Value);

                Assert.AreEqual(name ?? "Calibri", font.FontName.Val.Value);
                Assert.IsTrue(cellFormat.ApplyFont.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow(14D)]
        [DataRow(null)]
        public void SetFontTestSize(double? size)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font() { FontSize = new FontSize() { Val = 11D } }), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                var range = new Range(worksheet, 1, 1);

                range.SetFont(size: size);

                var cell = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cellFormat = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell.StyleIndex.Value);
                var font = stylesheet.GetFirstChild<Fonts>().Elements<Font>().ElementAt((int)cellFormat.FontId.Value);

                Assert.AreEqual(size == null ? 11D : size.Value, font.FontSize.Val.Value);
                Assert.IsTrue(cellFormat.ApplyFont.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow(true)]
        [DataRow(false)]
        [DataRow(null)]
        public void SetFontTestBold(bool? apply)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }, new Cell() { CellReference = "B1", StyleIndex = 1U }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font(), new Font(new Bold())), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U }, new CellFormat() { NumberFormatId = 0U, FontId = 1U, FillId = 0U, BorderId = 0U }));
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                var range = new Range(worksheet, 1, 1, 1, 2);

                range.SetFont(bold: apply);

                var cell1 = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cell2 = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().Elements<Cell>().Last();
                var cell1Format = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell1.StyleIndex.Value);
                var cell2Format = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell2.StyleIndex.Value);
                var font1 = stylesheet.GetFirstChild<Fonts>().Elements<Font>().ElementAt((int)cell1Format.FontId.Value);
                var font2 = stylesheet.GetFirstChild<Fonts>().Elements<Font>().ElementAt((int)cell2Format.FontId.Value);

                if (apply == null)
                {
                    Assert.IsFalse(font1.Elements<Bold>().Any());
                    Assert.IsTrue(font2.Elements<Bold>().Any());
                }
                else
                {
                    Assert.AreEqual(apply.Value, font1.Elements<Bold>().Any());
                    Assert.AreEqual(apply.Value, font2.Elements<Bold>().Any());
                }

                Assert.IsTrue(cell1Format.ApplyFont.Value);
                Assert.IsTrue(cell2Format.ApplyFont.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow(true)]
        [DataRow(false)]
        [DataRow(null)]
        public void SetFontTestUnderline(bool? apply)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }, new Cell() { CellReference = "B1", StyleIndex = 1U }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font(), new Font(new Underline())), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U }, new CellFormat() { NumberFormatId = 0U, FontId = 1U, FillId = 0U, BorderId = 0U }));
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                var range = new Range(worksheet, 1, 1, 1, 2);

                range.SetFont(underline: apply);

                var cell1 = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cell2 = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().Elements<Cell>().Last();
                var cell1Format = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell1.StyleIndex.Value);
                var cell2Format = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell2.StyleIndex.Value);
                var font1 = stylesheet.GetFirstChild<Fonts>().Elements<Font>().ElementAt((int)cell1Format.FontId.Value);
                var font2 = stylesheet.GetFirstChild<Fonts>().Elements<Font>().ElementAt((int)cell2Format.FontId.Value);

                if (apply == null)
                {
                    Assert.IsFalse(font1.Elements<Underline>().Any());
                    Assert.IsTrue(font2.Elements<Underline>().Any());
                }
                else
                {
                    Assert.AreEqual(apply.Value, font1.Elements<Underline>().Any());
                    Assert.AreEqual(apply.Value, font2.Elements<Underline>().Any());
                }

                Assert.IsTrue(cell1Format.ApplyFont.Value);
                Assert.IsTrue(cell2Format.ApplyFont.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow(true)]
        [DataRow(false)]
        [DataRow(null)]
        public void SetFontTestItalic(bool? apply)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }, new Cell() { CellReference = "B1", StyleIndex = 1U }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font(), new Font(new Italic())), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U }, new CellFormat() { NumberFormatId = 0U, FontId = 1U, FillId = 0U, BorderId = 0U }));
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                var range = new Range(worksheet, 1, 1, 1, 2);

                range.SetFont(italic: apply);

                var cell1 = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cell2 = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().Elements<Cell>().Last();
                var cell1Format = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell1.StyleIndex.Value);
                var cell2Format = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell2.StyleIndex.Value);
                var font1 = stylesheet.GetFirstChild<Fonts>().Elements<Font>().ElementAt((int)cell1Format.FontId.Value);
                var font2 = stylesheet.GetFirstChild<Fonts>().Elements<Font>().ElementAt((int)cell2Format.FontId.Value);

                if (apply == null)
                {
                    Assert.IsFalse(font1.Elements<Italic>().Any());
                    Assert.IsTrue(font2.Elements<Italic>().Any());
                }
                else
                {
                    Assert.AreEqual(apply.Value, font1.Elements<Italic>().Any());
                    Assert.AreEqual(apply.Value, font2.Elements<Italic>().Any());
                }

                Assert.IsTrue(cell1Format.ApplyFont.Value);
                Assert.IsTrue(cell2Format.ApplyFont.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow(true)]
        [DataRow(false)]
        [DataRow(null)]
        public void SetFontTestStrikethrough(bool? apply)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }, new Cell() { CellReference = "B1", StyleIndex = 1U }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font(), new Font(new Strike())), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U }, new CellFormat() { NumberFormatId = 0U, FontId = 1U, FillId = 0U, BorderId = 0U }));
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                var range = new Range(worksheet, 1, 1, 1, 2);

                range.SetFont(strikethrough: apply);

                var cell1 = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cell2 = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().Elements<Cell>().Last();
                var cell1Format = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell1.StyleIndex.Value);
                var cell2Format = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell2.StyleIndex.Value);
                var font1 = stylesheet.GetFirstChild<Fonts>().Elements<Font>().ElementAt((int)cell1Format.FontId.Value);
                var font2 = stylesheet.GetFirstChild<Fonts>().Elements<Font>().ElementAt((int)cell2Format.FontId.Value);

                if (apply == null)
                {
                    Assert.IsFalse(font1.Elements<Strike>().Any());
                    Assert.IsTrue(font2.Elements<Strike>().Any());
                }
                else
                {
                    Assert.AreEqual(apply.Value, font1.Elements<Strike>().Any());
                    Assert.AreEqual(apply.Value, font2.Elements<Strike>().Any());
                }

                Assert.IsTrue(cell1Format.ApplyFont.Value);
                Assert.IsTrue(cell2Format.ApplyFont.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow("FF0000FF")]
        [DataRow(null)]
        public void SetFontTestColor(string color)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font() { Color = new Color() { Rgb = "FFFF0000" } }), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                var range = new Range(worksheet, 1, 1);

                range.SetFont(color: color);

                var cell = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cellFormat = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell.StyleIndex.Value);
                var font = stylesheet.GetFirstChild<Fonts>().Elements<Font>().ElementAt((int)cellFormat.FontId.Value);

                Assert.AreEqual(color ?? "FFFF0000", font.Color.Rgb.Value);
                Assert.IsTrue(cellFormat.ApplyFont.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow(GradientFillType.Vertical)]
        [DataRow(GradientFillType.VerticalCenter)]
        [DataRow(GradientFillType.Horizontal)]
        [DataRow(GradientFillType.HorizontalCenter)]
        public void SetGradientFillTest(GradientFillType fillType)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                var range = new Range(worksheet, 1, 1);
                var fillColor1 = "FFFF0000";
                var fillColor2 = "FF0000FF";

                range.SetGradientFill(fillType, fillColor1, fillColor2);

                var cell = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cellFormat = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell.StyleIndex.Value);
                var fill = stylesheet.GetFirstChild<Fills>().Elements<Fill>().ElementAt((int)cellFormat.FillId.Value);

                var expectedFill = new Fill() { GradientFill = new GradientFill() { Type = GradientValues.Linear } };
                switch(fillType)
                {
                    case GradientFillType.Vertical:
                        expectedFill.GradientFill.Degree = 90D;
                        expectedFill.GradientFill.Append(new GradientStop(new Color() { Rgb = fillColor1 }) { Position = 0D });
                        expectedFill.GradientFill.Append(new GradientStop(new Color() { Rgb = fillColor2 }) { Position = 1D });
                        break;

                    case GradientFillType.VerticalCenter:
                        expectedFill.GradientFill.Degree = 90D;
                        expectedFill.GradientFill.Append(new GradientStop(new Color() { Rgb = fillColor1 }) { Position = 0D });
                        expectedFill.GradientFill.Append(new GradientStop(new Color() { Rgb = fillColor2 }) { Position = 0.5D });
                        expectedFill.GradientFill.Append(new GradientStop(new Color() { Rgb = fillColor1 }) { Position = 1D });
                        break;

                    case GradientFillType.Horizontal:
                        expectedFill.GradientFill.Degree = 0D;
                        expectedFill.GradientFill.Append(new GradientStop(new Color() { Rgb = fillColor1 }) { Position = 0D });
                        expectedFill.GradientFill.Append(new GradientStop(new Color() { Rgb = fillColor2 }) { Position = 1D });
                        break;

                    case GradientFillType.HorizontalCenter:
                        expectedFill.GradientFill.Degree = 0D;
                        expectedFill.GradientFill.Append(new GradientStop(new Color() { Rgb = fillColor1 }) { Position = 0D });
                        expectedFill.GradientFill.Append(new GradientStop(new Color() { Rgb = fillColor2 }) { Position = 0.5D });
                        expectedFill.GradientFill.Append(new GradientStop(new Color() { Rgb = fillColor1 }) { Position = 1D });
                        break;
                }

                Assert.AreEqual(expectedFill.OuterXml, fill.OuterXml);
                Assert.IsTrue(cellFormat.ApplyFill.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow(PatternValues.DarkDown, "FF000000", null)]
        [DataRow(PatternValues.DarkGray, null, "FF000000")]
        public void SetPatternFillTest(PatternValues patternType, string foregroundColor, string backgroundColor)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                var range = new Range(worksheet, 1, 1);

                range.SetPatternFill(patternType, foregroundColor, backgroundColor);

                var cell = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cellFormat = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell.StyleIndex.Value);
                var fill = stylesheet.GetFirstChild<Fills>().Elements<Fill>().ElementAt((int)cellFormat.FillId.Value);

                var expectedFill = new Fill() { PatternFill = new PatternFill() { PatternType = patternType } };
                if (foregroundColor != null)
                {
                    expectedFill.PatternFill.Append(new ForegroundColor() { Rgb = foregroundColor });
                }
                if (backgroundColor != null)
                {
                    expectedFill.PatternFill.Append(new BackgroundColor() { Rgb = backgroundColor });
                }

                Assert.AreEqual(expectedFill.OuterXml, fill.OuterXml);
                Assert.IsTrue(cellFormat.ApplyFill.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow(null, "FF000001", BorderStyleValues.DashDotDot, "FF000002", BorderStyleValues.Dashed, "FF000003", BorderStyleValues.Dotted, "FF000004", BorderStyleValues.Double, "FF000005", true, false)]
        [DataRow(BorderStyleValues.DashDot, null, BorderStyleValues.DashDotDot, "FF000002", BorderStyleValues.Dashed, "FF000003", BorderStyleValues.Dotted, "FF000004", BorderStyleValues.Double, "FF000005", true, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", null, "FF000002", BorderStyleValues.Dashed, "FF000003", BorderStyleValues.Dotted, "FF000004", BorderStyleValues.Double, "FF000005", true, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", BorderStyleValues.DashDotDot, null, BorderStyleValues.Dashed, "FF000003", BorderStyleValues.Dotted, "FF000004", BorderStyleValues.Double, "FF000005", true, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", BorderStyleValues.DashDotDot, "FF000002", null, "FF000003", BorderStyleValues.Dotted, "FF000004", BorderStyleValues.Double, "FF000005", true, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", BorderStyleValues.DashDotDot, "FF000002", BorderStyleValues.Dashed, null, BorderStyleValues.Dotted, "FF000004", BorderStyleValues.Double, "FF000005", true, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", BorderStyleValues.DashDotDot, "FF000002", BorderStyleValues.Dashed, "FF000003", null, "FF000004", BorderStyleValues.Double, "FF000005", true, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", BorderStyleValues.DashDotDot, "FF000002", BorderStyleValues.Dashed, "FF000003", BorderStyleValues.Dotted, null, BorderStyleValues.Double, "FF000005", true, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", BorderStyleValues.DashDotDot, "FF000002", BorderStyleValues.Dashed, "FF000003", BorderStyleValues.Dotted, "FF000004", null, "FF000005", true, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", BorderStyleValues.DashDotDot, "FF000002", BorderStyleValues.Dashed, "FF000003", BorderStyleValues.Dotted, "FF000004", BorderStyleValues.Double, null, true, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", BorderStyleValues.DashDotDot, "FF000002", BorderStyleValues.Dashed, "FF000003", BorderStyleValues.Dotted, "FF000004", BorderStyleValues.Double, "FF000005", null, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", BorderStyleValues.DashDotDot, "FF000002", BorderStyleValues.Dashed, "FF000003", BorderStyleValues.Dotted, "FF000004", BorderStyleValues.Double, "FF000005", true, null)]
        public void SetBorderTest(BorderStyleValues? leftBorderStyle, string leftBorderColor, BorderStyleValues? rightBorderStyle, string rightBorderColor, BorderStyleValues? topBorderStyle, string topBorderColor, BorderStyleValues? bottomBorderStyle, string bottomBorderColor, BorderStyleValues? diagonalBorderStyle, string diagonalBorderColor, bool? diagonalUp, bool? diagonalDown)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                var range = new Range(worksheet, 1, 1);

                range.SetBorder(leftBorderStyle: leftBorderStyle, leftBorderColor: leftBorderColor,
                    rightBorderStyle: rightBorderStyle, rightBorderColor: rightBorderColor,
                    topBorderStyle: topBorderStyle, topBorderColor: topBorderColor,
                    bottomBorderStyle: bottomBorderStyle, bottomBorderColor: bottomBorderColor,
                    diagonalBorderStyle: diagonalBorderStyle, diagonalBorderColor: diagonalBorderColor,
                    diagonalUp: diagonalUp, diagonalDown: diagonalDown);

                var cell = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cellFormat = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell.StyleIndex.Value);
                var border = stylesheet.GetFirstChild<Borders>().Elements<Border>().ElementAt((int)cellFormat.BorderId.Value);

                var expectedBorder = new Border();
                if (leftBorderStyle != null)
                {
                    expectedBorder.LeftBorder ??= new LeftBorder();
                    expectedBorder.LeftBorder.Style = leftBorderStyle;
                }
                if (leftBorderColor != null)
                {
                    expectedBorder.LeftBorder ??= new LeftBorder();
                    expectedBorder.LeftBorder.Color = new Color() { Rgb = leftBorderColor };
                }
                if (rightBorderStyle != null)
                {
                    expectedBorder.RightBorder ??= new RightBorder();
                    expectedBorder.RightBorder.Style = rightBorderStyle;
                }
                if (rightBorderColor != null)
                {
                    expectedBorder.RightBorder ??= new RightBorder();
                    expectedBorder.RightBorder.Color = new Color() { Rgb = rightBorderColor };
                }
                if (topBorderStyle != null)
                {
                    expectedBorder.TopBorder ??= new TopBorder();
                    expectedBorder.TopBorder.Style = topBorderStyle;
                }
                if (topBorderColor != null)
                {
                    expectedBorder.TopBorder ??= new TopBorder();
                    expectedBorder.TopBorder.Color = new Color() { Rgb = topBorderColor };
                }
                if (bottomBorderStyle != null)
                {
                    expectedBorder.BottomBorder ??= new BottomBorder();
                    expectedBorder.BottomBorder.Style = bottomBorderStyle;
                }
                if (bottomBorderColor != null)
                {
                    expectedBorder.BottomBorder ??= new BottomBorder();
                    expectedBorder.BottomBorder.Color = new Color() { Rgb = bottomBorderColor };
                }
                if (diagonalBorderStyle != null)
                {
                    expectedBorder.DiagonalBorder ??= new DiagonalBorder();
                    expectedBorder.DiagonalBorder.Style = diagonalBorderStyle;
                }
                if (diagonalBorderColor != null)
                {
                    expectedBorder.DiagonalBorder ??= new DiagonalBorder();
                    expectedBorder.DiagonalBorder.Color = new Color() { Rgb = diagonalBorderColor };
                }
                if (diagonalUp != null)
                {
                    expectedBorder.DiagonalUp = diagonalUp.Value;
                }
                if (diagonalDown != null)
                {
                    expectedBorder.DiagonalDown = diagonalDown.Value;
                }

                Assert.AreEqual(expectedBorder.OuterXml, border.OuterXml);
                Assert.IsTrue(cellFormat.ApplyBorder.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow(null, "FF000001", BorderStyleValues.DashDotDot, "FF000002", true, false)]
        [DataRow(BorderStyleValues.DashDot, null, BorderStyleValues.DashDotDot, "FF000002", true, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", null, "FF000002", true, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", BorderStyleValues.DashDotDot, null, true, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", BorderStyleValues.DashDotDot, "FF000002", null, false)]
        [DataRow(BorderStyleValues.DashDot, "FF000001", BorderStyleValues.DashDotDot, "FF000002", true, null)]
        public void SetOutlineBorderTest(BorderStyleValues? outlineBorderStyle, string outlineBorderColor, BorderStyleValues? diagonalBorderStyle, string diagonalBorderColor, bool? diagonalUp, bool? diagonalDown)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                var range = new Range(worksheet, 1, 1);

                range.SetOutlineBorder(outlineBorderStyle: outlineBorderStyle, outlineBorderColor: outlineBorderColor,
                    diagonalBorderStyle: diagonalBorderStyle, diagonalBorderColor: diagonalBorderColor,
                    diagonalUp: diagonalUp, diagonalDown: diagonalDown);

                var cell = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cellFormat = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell.StyleIndex.Value);
                var border = stylesheet.GetFirstChild<Borders>().Elements<Border>().ElementAt((int)cellFormat.BorderId.Value);

                var expectedBorder = new Border();
                if (outlineBorderStyle != null)
                {
                    expectedBorder.LeftBorder ??= new LeftBorder();
                    expectedBorder.LeftBorder.Style = outlineBorderStyle;
                    expectedBorder.RightBorder ??= new RightBorder();
                    expectedBorder.RightBorder.Style = outlineBorderStyle;
                    expectedBorder.TopBorder ??= new TopBorder();
                    expectedBorder.TopBorder.Style = outlineBorderStyle;
                    expectedBorder.BottomBorder ??= new BottomBorder();
                    expectedBorder.BottomBorder.Style = outlineBorderStyle;
                }
                if (outlineBorderColor != null)
                {
                    expectedBorder.LeftBorder ??= new LeftBorder();
                    expectedBorder.LeftBorder.Color = new Color() { Rgb = outlineBorderColor };
                    expectedBorder.RightBorder ??= new RightBorder();
                    expectedBorder.RightBorder.Color = new Color() { Rgb = outlineBorderColor };
                    expectedBorder.TopBorder ??= new TopBorder();
                    expectedBorder.TopBorder.Color = new Color() { Rgb = outlineBorderColor };
                    expectedBorder.BottomBorder ??= new BottomBorder();
                    expectedBorder.BottomBorder.Color = new Color() { Rgb = outlineBorderColor };
                }
                if (diagonalBorderStyle != null)
                {
                    expectedBorder.DiagonalBorder ??= new DiagonalBorder();
                    expectedBorder.DiagonalBorder.Style = diagonalBorderStyle;
                }
                if (diagonalBorderColor != null)
                {
                    expectedBorder.DiagonalBorder ??= new DiagonalBorder();
                    expectedBorder.DiagonalBorder.Color = new Color() { Rgb = diagonalBorderColor };
                }
                if (diagonalUp != null)
                {
                    expectedBorder.DiagonalUp = diagonalUp.Value;
                }
                if (diagonalDown != null)
                {
                    expectedBorder.DiagonalDown = diagonalDown.Value;
                }

                Assert.AreEqual(expectedBorder.OuterXml, border.OuterXml);
                Assert.IsTrue(cellFormat.ApplyBorder.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow(null, VerticalAlignmentValues.Bottom, 1, true)]
        [DataRow(HorizontalAlignmentValues.Center, null, 1, true)]
        [DataRow(HorizontalAlignmentValues.Center, VerticalAlignmentValues.Bottom, null, true)]
        [DataRow(HorizontalAlignmentValues.Center, VerticalAlignmentValues.Bottom, 1, null)]
        [DataRow(HorizontalAlignmentValues.CenterContinuous, VerticalAlignmentValues.Center, 2, false)]
        public void SetAlignmentTest(HorizontalAlignmentValues? horizontalAlignment, VerticalAlignmentValues? verticalAlignment, int? indent, bool? wrap)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                var range = new Range(worksheet, 1, 1);

                range.SetAlignment(horizontalAlignment, verticalAlignment, indent, wrap);

                var cell = worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().GetFirstChild<Cell>();
                var cellFormat = stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)cell.StyleIndex.Value);

                var expectedCellFormat = new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyAlignment = true };
                if (horizontalAlignment != null)
                {
                    expectedCellFormat.Alignment ??= new Alignment();
                    expectedCellFormat.Alignment.Horizontal = horizontalAlignment.Value;
                }
                if (verticalAlignment != null)
                {
                    expectedCellFormat.Alignment ??= new Alignment();
                    expectedCellFormat.Alignment.Vertical = verticalAlignment.Value;
                }
                if (indent != null)
                {
                    expectedCellFormat.Alignment ??= new Alignment();
                    expectedCellFormat.Alignment.Indent = (uint)indent.Value;
                }
                if (wrap != null)
                {
                    expectedCellFormat.Alignment ??= new Alignment();
                    expectedCellFormat.Alignment.WrapText = wrap.Value;
                }

                Assert.AreEqual(expectedCellFormat.OuterXml, cellFormat.OuterXml);
                Assert.IsTrue(cellFormat.ApplyAlignment.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow(1.0D)]
        [DataRow(1.0F)]
        [DataRow(1)]
        [DataRow("string")]
        public void SetValuesTestSingleValue(object value)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }, new Cell() { CellReference = "B1" }) { RowIndex = 1U },
                    new Row(new Cell() { CellReference = "A2" }, new Cell() { CellReference = "B2" }) { RowIndex = 2U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var range = new Range(worksheet, 1, 1, 2, 2);

                range.SetValues(value);

                CellValues expectedValueType;
                CellValue expectedValue;
                if (value is string)
                {
                    expectedValueType = CellValues.SharedString;
                    expectedValue = new CellValue(0);
                }
                else
                {
                    expectedValueType = CellValues.Number;
                    expectedValue = new CellValue((decimal)(dynamic)value);
                }

                Assert.IsTrue(worksheet.Descendants<Cell>().All(c => c.DataType.Value == expectedValueType && c.CellValue.InnerText == expectedValue.InnerText));
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void SetValuesTestOneDimensionalArray()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }, new Cell() { CellReference = "B1" }, new Cell() { CellReference = "C1" }, new Cell() { CellReference = "D1" }) { RowIndex = 1U },
                    new Row(new Cell() { CellReference = "A2" }, new Cell() { CellReference = "B2" }, new Cell() { CellReference = "C2" }, new Cell() { CellReference = "D2" }) { RowIndex = 2U },
                    new Row(new Cell() { CellReference = "A3" }, new Cell() { CellReference = "B3" }, new Cell() { CellReference = "C3" }, new Cell() { CellReference = "D3" }) { RowIndex = 3U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var cellRange = new Range(worksheet, 1, 1, 2, 3);
                var rowRange = new Range(worksheet, 3, null);
                var columnRange = new Range(worksheet, null, 4);
                var values = new object[] { 1D, 1F, "string" };

                cellRange.SetValues(values);
                rowRange.SetValues(values);
                columnRange.SetValues(values);

                Func<string, object, bool> CellIsValid = delegate (string cellReference, object value)
                {
                    var cell = worksheet.Descendants<Cell>().Single(c => c.CellReference.Value == cellReference);

                    CellValues expectedValueType;
                    CellValue expectedValue;
                    if (value is string)
                    {
                        expectedValueType = CellValues.SharedString;
                        expectedValue = new CellValue(document.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().Single(s => s.InnerText.Equals(value)).ElementsBefore().Count());
                    }
                    else
                    {
                        expectedValueType = CellValues.Number;
                        expectedValue = new CellValue((decimal)(dynamic)value);
                    }

                    return cell.DataType.Value == expectedValueType && cell.CellValue.InnerText == expectedValue.InnerText;
                };
                
                Assert.IsTrue(CellIsValid("A1", values[0]));
                Assert.IsTrue(CellIsValid("B1", values[1]));
                Assert.IsTrue(CellIsValid("C1", values[2]));
                Assert.IsTrue(CellIsValid("A2", values[0]));
                Assert.IsTrue(CellIsValid("B2", values[1]));
                Assert.IsTrue(CellIsValid("C2", values[2]));

                Assert.IsTrue(CellIsValid("A3", values[0]));
                Assert.IsTrue(CellIsValid("B3", values[1]));
                Assert.IsTrue(CellIsValid("C3", values[2]));

                Assert.IsTrue(CellIsValid("D1", values[0]));
                Assert.IsTrue(CellIsValid("D2", values[1]));
                Assert.IsTrue(CellIsValid("D3", values[2]));
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void SetValuesTestTwoDimensionalArray()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }, new Cell() { CellReference = "B1" }, new Cell() { CellReference = "C1" }, new Cell() { CellReference = "D1" }) { RowIndex = 1U },
                    new Row(new Cell() { CellReference = "A2" }, new Cell() { CellReference = "B2" }, new Cell() { CellReference = "C2" }, new Cell() { CellReference = "D2" }) { RowIndex = 2U },
                    new Row(new Cell() { CellReference = "A3" }, new Cell() { CellReference = "B3" }, new Cell() { CellReference = "C3" }, new Cell() { CellReference = "D3" }) { RowIndex = 3U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var cellRange = new Range(worksheet, 1, 1, 2, 3);
                var rowRange = new Range(worksheet, 3, null);
                var columnRange = new Range(worksheet, null, 4);
                var values = new object[,] { { 1D, 1F, "string" }, { 2U, 3F, "other string" } };

                cellRange.SetValues(values);
                rowRange.SetValues(values);
                columnRange.SetValues(values);

                Func<string, object, bool> CellIsValid = delegate (string cellReference, object value)
                {
                    var cell = worksheet.Descendants<Cell>().Single(c => c.CellReference.Value == cellReference);

                    CellValues expectedValueType;
                    CellValue expectedValue;
                    if (value is string)
                    {
                        expectedValueType = CellValues.SharedString;
                        expectedValue = new CellValue(document.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().Single(s => s.InnerText.Equals(value)).ElementsBefore().Count());
                    }
                    else
                    {
                        expectedValueType = CellValues.Number;
                        expectedValue = new CellValue((decimal)(dynamic)value);
                    }

                    return cell.DataType.Value == expectedValueType && cell.CellValue.InnerText == expectedValue.InnerText;
                };
                
                Assert.IsTrue(CellIsValid("A1", values[0, 0]));
                Assert.IsTrue(CellIsValid("B1", values[0, 1]));
                Assert.IsTrue(CellIsValid("C1", values[0, 2]));
                Assert.IsTrue(CellIsValid("A2", values[1, 0]));
                Assert.IsTrue(CellIsValid("B2", values[1, 1]));
                Assert.IsTrue(CellIsValid("C2", values[1, 2]));

                Assert.IsTrue(CellIsValid("A3", values[0, 0]));
                Assert.IsTrue(CellIsValid("B3", values[0, 1]));
                Assert.IsTrue(CellIsValid("C3", values[0, 2]));

                Assert.IsTrue(CellIsValid("D1", values[0, 0]));
                Assert.IsTrue(CellIsValid("D2", values[0, 1]));
                Assert.IsTrue(CellIsValid("D3", values[0, 2]));
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }
    }
}