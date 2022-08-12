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
    public class SpreadsheetUtilityTests
    {
        [TestMethod()]
        public void GetColumnLetterTestConvertNumber()
        {
            Assert.AreEqual("A", SpreadsheetUtility.GetColumnLetter(1));
            Assert.AreEqual("Z", SpreadsheetUtility.GetColumnLetter(26));
            Assert.AreEqual("AA", SpreadsheetUtility.GetColumnLetter(27));
            Assert.AreEqual("AAA", SpreadsheetUtility.GetColumnLetter(26 * 27 + 1));
        }

        [TestMethod()]
        public void GetColumnLetterTestString()
        {
            Assert.AreEqual("A", SpreadsheetUtility.GetColumnLetter("A1"));
            Assert.AreEqual("ZZ", SpreadsheetUtility.GetColumnLetter("ZZ1"));
        }

        [TestMethod()]
        public void GetRowIndexTest()
        {
            Assert.AreEqual(1U, SpreadsheetUtility.GetRowIndex("A1"));
            Assert.AreEqual(2U, SpreadsheetUtility.GetRowIndex("ZZ2"));
        }
    }

    [TestClass()]
    public class SpreadsheetDocumentExtensionTest
    {
        [TestMethod()]
        [DataRow(BuiltinCellStyle.Normal, "Normal")]
        [DataRow(BuiltinCellStyle.Hyperlink, "Hyperlink")]
        public void AddCellStyleTest(BuiltinCellStyle builtinCellStyle, string name)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook();
                var stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                
                document.AddCellStyle(builtinCellStyle);

                var cellStyle = stylesheet.Elements<CellStyles>().Single().Elements<CellStyle>().SingleOrDefault(s => s.BuiltinId.Value == (uint)builtinCellStyle);

                Assert.IsNotNull(cellStyle);
                Assert.AreEqual(name, cellStyle.Name.Value);
                Assert.IsTrue(stylesheet.Elements<CellStyleFormats>().Single().Elements<CellFormat>().Count() > cellStyle.FormatId.Value);
                Assert.AreEqual(1U, stylesheet.Elements<CellStyles>().Single().Count.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow(BuiltinCellStyle.Normal)]
        [DataRow(BuiltinCellStyle.Hyperlink)]
        public void ApplyCellStyleTest(BuiltinCellStyle builtinCellStyle)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook();
                var stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;
                var inputCellFormat = new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U };
                var outputCellFormat = inputCellFormat.Clone() as CellFormat;

                document.ApplyCellStyle(outputCellFormat, builtinCellStyle);

                var cellStyle = stylesheet.Elements<CellStyles>().Single().Elements<CellStyle>().Single(s => s.BuiltinId == (uint)builtinCellStyle);
                var styleFormat = stylesheet.Elements<CellStyleFormats>().Single().Elements<CellFormat>().ElementAt((int)cellStyle.FormatId.Value);
                var expectedFormat = new CellFormat();

                expectedFormat.NumberFormatId = (styleFormat.ApplyNumberFormat?.Value ?? true) ? styleFormat.NumberFormatId.Value : inputCellFormat.NumberFormatId.Value;
                expectedFormat.FontId = (styleFormat.ApplyFont?.Value ?? true) ? styleFormat.FontId.Value : inputCellFormat.FontId.Value;
                expectedFormat.FillId = (styleFormat.ApplyFill?.Value ?? true) ? styleFormat.FillId.Value : inputCellFormat.FillId.Value;
                expectedFormat.BorderId = (styleFormat.ApplyBorder?.Value ?? true) ? styleFormat.BorderId.Value : inputCellFormat.BorderId.Value;
                expectedFormat.Alignment = ((styleFormat.ApplyAlignment?.Value ?? true) ? styleFormat.Alignment?.Clone() : inputCellFormat.Alignment?.Clone()) as Alignment;
                expectedFormat.Protection = ((styleFormat.ApplyProtection?.Value ?? true) ? styleFormat.Protection?.Clone() : inputCellFormat.Protection?.Clone()) as Protection;

                Assert.AreEqual(expectedFormat.NumberFormatId.Value, outputCellFormat.NumberFormatId.Value);
                Assert.AreEqual(expectedFormat.FontId.Value, outputCellFormat.FontId.Value);
                Assert.AreEqual(expectedFormat.FillId.Value, outputCellFormat.FillId.Value);
                Assert.AreEqual(expectedFormat.BorderId.Value, outputCellFormat.BorderId.Value);
                Assert.AreEqual(expectedFormat.Alignment?.OuterXml, outputCellFormat.Alignment?.OuterXml);
                Assert.AreEqual(expectedFormat.Protection?.OuterXml, outputCellFormat.Protection?.OuterXml);
                
                Assert.AreEqual(outputCellFormat.FormatId.Value, cellStyle.FormatId.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void AddNewStylesheetTest()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook();
                
                document.AddNewStylesheet();

                var stylesheet = document.WorkbookPart.WorkbookStylesPart.Stylesheet;
                var font = stylesheet.Elements<Fonts>().Single().Elements<Font>().ElementAt(0);
                var fillArray = stylesheet.Elements<Fills>().Single().Elements<Fill>().ToArray();

                Assert.AreEqual("Calibri", font.FontName.Val.Value);
                Assert.AreEqual(11D, font.FontSize.Val.Value);
                Assert.AreEqual(PatternValues.None, fillArray[0].PatternFill.PatternType.Value);
                Assert.AreEqual(PatternValues.Gray125, fillArray[1].PatternFill.PatternType.Value);
                Assert.AreEqual(1, stylesheet.Elements<Borders>().Single().ChildElements.Count);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void AddNewStylesheetTestWithParams()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook();
                
                document.AddNewStylesheet(new Font() { FontName = new FontName() { Val = "Helvetica" }, FontSize = new FontSize() { Val = 14D } },
                    new Fill() { PatternFill = new PatternFill() { PatternType = PatternValues.DarkHorizontal } },
                    new Border() { LeftBorder = new LeftBorder() { Style = BorderStyleValues.Hair } });

                var stylesheet = document.WorkbookPart.WorkbookStylesPart.Stylesheet;
                var font = stylesheet.Elements<Fonts>().Single().Elements<Font>().ElementAt(0);
                var fillArray = stylesheet.Elements<Fills>().Single().Elements<Fill>().ToArray();
                var border = stylesheet.Elements<Borders>().Single().Elements<Border>().ElementAt(1);

                Assert.AreEqual("Helvetica", font.FontName.Val.Value);
                Assert.AreEqual(14D, font.FontSize.Val.Value);
                Assert.AreEqual(PatternValues.None, fillArray[0].PatternFill.PatternType.Value);
                Assert.AreEqual(PatternValues.Gray125, fillArray[1].PatternFill.PatternType.Value);
                Assert.AreEqual(PatternValues.DarkHorizontal, fillArray[2].PatternFill.PatternType.Value);
                Assert.AreEqual(BorderStyleValues.Hair, border.LeftBorder.Style.Value);
                Assert.AreEqual(1, stylesheet.Elements<CellStyleFormats>().Single().ChildElements.Count);
                Assert.AreEqual(1, stylesheet.Elements<CellFormats>().Single().ChildElements.Count);
                Assert.IsNotNull(stylesheet.Elements<CellStyles>().Single().Elements<CellStyle>().SingleOrDefault(s => s.Name.Value == "Normal"));
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void CreateStyleElementTemplateTestNullId()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook();
                //Create stylesheet with 2 different Fonts, Fills, and Borders, with a CellFormat pointing to each set.
                var stylesheet = new Stylesheet(
                    new Fonts(new Font() { FontName = new FontName() { Val = "Calibri" }, FontSize = new FontSize() { Val = 11D } }, new Font() { FontName = new FontName() { Val = "Helvetica" }, FontSize = new FontSize() { Val = 14D } }),
                    new Fills(new Fill() { PatternFill = new PatternFill() { PatternType = PatternValues.None } }, new Fill() { PatternFill = new PatternFill() { PatternType = PatternValues.DarkHorizontal } }),
                    new Borders(new Border(), new Border() { LeftBorder = new LeftBorder() { Style = BorderStyleValues.Hair } }),
                    new CellStyleFormats(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U }),
                    new CellFormats(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U }, new CellFormat() { NumberFormatId = 1U, FontId = 1U, FillId = 1U, BorderId = 1U, FormatId = 0U }),
                    new CellStyles(new CellStyle() { BuiltinId = 0U, Name = "Normal", FormatId = 0U }));
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;

                var cellFormat = document.CreateStyleElementTemplate<CellFormat>(null);
                var font = document.CreateStyleElementTemplate<Font>(null);
                var fill = document.CreateStyleElementTemplate<Fill>(null);
                var border = document.CreateStyleElementTemplate<Border>(null);

                var normalStyle = stylesheet.Elements<CellStyles>().Single().Elements<CellStyle>().SingleOrDefault(s => s.BuiltinId == 0U);
                var normalStyleFormat = stylesheet.Elements<CellStyleFormats>().Single().Elements<CellFormat>().ElementAt((int)normalStyle.FormatId.Value);

                var fontArray = stylesheet.Elements<Fonts>().Single().Elements<Font>().ToArray();
                var fillArray = stylesheet.Elements<Fills>().Single().Elements<Fill>().ToArray();
                var borderArray = stylesheet.Elements<Borders>().Single().Elements<Border>().ToArray();

                var expectedFont = fontArray[normalStyleFormat.FontId.Value];
                var expectedFill = fillArray[normalStyleFormat.FillId.Value];
                var expectedBorder = borderArray[normalStyleFormat.FillId.Value];

                Assert.AreEqual(normalStyleFormat.NumberFormatId.Value, cellFormat.NumberFormatId.Value);
                Assert.AreEqual(normalStyleFormat.FontId.Value, cellFormat.FontId.Value);
                Assert.AreEqual(normalStyleFormat.FillId.Value, cellFormat.FillId.Value);
                Assert.AreEqual(normalStyleFormat.BorderId.Value, cellFormat.BorderId.Value);
                Assert.AreEqual(normalStyle.FormatId.Value, cellFormat.FormatId.Value);

                Assert.AreEqual(expectedFont.OuterXml, font.OuterXml);
                Assert.AreEqual(expectedFill.OuterXml, fill.OuterXml);
                Assert.AreEqual(expectedBorder.OuterXml, border.OuterXml);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        [DataRow(0U)]
        [DataRow(1U)]
        public void CreateStyleElementTemplateTestNonNullId(uint rawStyleId)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook();
                //Create stylesheet with 2 different Fonts, Fills, and Borders, with a CellFormat pointing to each set.
                var stylesheet = new Stylesheet(
                    new Fonts(new Font() { FontName = new FontName() { Val = "Calibri" }, FontSize = new FontSize() { Val = 11D } }, new Font() { FontName = new FontName() { Val = "Helvetica" }, FontSize = new FontSize() { Val = 14D } }),
                    new Fills(new Fill() { PatternFill = new PatternFill() { PatternType = PatternValues.None } }, new Fill() { PatternFill = new PatternFill() { PatternType = PatternValues.DarkHorizontal } }),
                    new Borders(new Border(), new Border() { LeftBorder = new LeftBorder() { Style = BorderStyleValues.Hair } }),
                    new CellStyleFormats(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U }),
                    new CellFormats(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U }, new CellFormat() { NumberFormatId = 1U, FontId = 1U, FillId = 1U, BorderId = 1U, FormatId = 0U }),
                    new CellStyleFormats(new CellStyle() { BuiltinId = 0U, Name = "Normal", FormatId = 0U }));
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;

                var styleId = new UInt32Value(rawStyleId);
                var cellFormat = document.CreateStyleElementTemplate<CellFormat>(styleId);
                var font = document.CreateStyleElementTemplate<Font>(styleId);
                var fill = document.CreateStyleElementTemplate<Fill>(styleId);
                var border = document.CreateStyleElementTemplate<Border>(styleId);

                var fontArray = stylesheet.Elements<Fonts>().Single().Elements<Font>().ToArray();
                var fillArray = stylesheet.Elements<Fills>().Single().Elements<Fill>().ToArray();
                var borderArray = stylesheet.Elements<Borders>().Single().Elements<Border>().ToArray();

                var expectedCellFormat = stylesheet.Elements<CellFormats>().Single().Elements<CellFormat>().ElementAt((int)styleId.Value).Clone() as CellFormat;
                var expectedFont = fontArray[expectedCellFormat.FontId.Value];
                var expectedFill = fillArray[expectedCellFormat.FillId.Value];
                var expectedBorder = borderArray[expectedCellFormat.FillId.Value];

                Assert.AreEqual(expectedCellFormat.OuterXml, cellFormat.OuterXml);
                Assert.AreEqual(expectedFont.OuterXml, font.OuterXml);
                Assert.AreEqual(expectedFill.OuterXml, fill.OuterXml);
                Assert.AreEqual(expectedBorder.OuterXml, border.OuterXml);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void PrepNewDocumentTest()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.PrepNewDocument();

                var sheets = document.WorkbookPart?.Workbook?.Elements<Sheets>().SingleOrDefault();
                Assert.IsNotNull(sheets);
                Assert.AreEqual(1, sheets.ChildElements.Count);
                Assert.IsInstanceOfType(document.WorkbookPart.GetPartById(sheets.GetFirstChild<Sheet>().Id), typeof(WorksheetPart));
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void RemoveSharedStringItemTestUsedString()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" } ));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1", DataType = CellValues.SharedString, CellValue = new CellValue(0) }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<SharedStringTablePart>().SharedStringTable = new SharedStringTable(new SharedStringItem(new Text("test string"))) { Count = 1U };

                document.RemoveSharedStringItem(0U);

                Assert.AreEqual(1, document.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.Count);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        public void RemoveSharedStringItemTestUnusedString()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1", DataType = CellValues.SharedString, CellValue = new CellValue(1) }, new Cell() { CellReference = "B1", DataType = CellValues.Number, CellValue = new CellValue(0) }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                document.WorkbookPart.AddNewPart<SharedStringTablePart>().SharedStringTable = new SharedStringTable(new SharedStringItem(new Text("unused string")), new SharedStringItem(new Text("used string"))) { Count = 2U };

                document.RemoveSharedStringItem(0U);

                var table = document.WorkbookPart.SharedStringTablePart.SharedStringTable;

                Assert.AreEqual(1, table.ChildElements.Count);
                Assert.AreEqual("used string", table.GetFirstChild<SharedStringItem>().Text.Text);
                Assert.AreEqual("0", worksheet.Descendants<Cell>().Single(c => c.DataType.Value == CellValues.SharedString).CellValue.Text);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void WorksheetsTest()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                Assert.IsInstanceOfType(document.Worksheets(), typeof(WorksheetCollection));
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void WorksheetsTestStringIndex()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }, new Sheet() { Id = "rId2", SheetId = 2U, Name = "OtherSheet" }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = new Worksheet(new SheetData());
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId2").Worksheet = new Worksheet();

                Assert.IsTrue(document.Worksheets("TestSheet").HasChildren);
                Assert.IsFalse(document.Worksheets("OtherSheet").HasChildren);
                Assert.IsNull(document.Worksheets("NonExistent"));
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void WorksheetsTestNumberIndex()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }, new Sheet() { Id = "rId2", SheetId = 2U, Name = "OtherSheet" }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = new Worksheet(new SheetData());
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId2").Worksheet = new Worksheet();

                Assert.IsTrue(document.Worksheets(0).HasChildren);
                Assert.IsFalse(document.Worksheets(1).HasChildren);
                Assert.ThrowsException<ArgumentOutOfRangeException>(() => document.Worksheets(2));
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }
    }

    [TestClass()]
    public class WorksheetExtensionTest
    {
        [TestMethod()]
        public void ColumnsTest()
        {
            Assert.IsInstanceOfType(new Worksheet(new SheetData()).Columns(), typeof(Columns));
            Assert.IsInstanceOfType(new Worksheet(new Columns(), new SheetData()).Columns(), typeof(Columns));
        }

        [TestMethod()]
        [DataRow(new uint[] { 1U, 2U })]
        [DataRow(new uint[] { 1U, 3U })]
        public void ColumnsTestMinMax(uint[] indices)
        {
            var worksheet = new Worksheet(new Columns(new Column() { Min = 1U, Max = 2U }), new SheetData());
            var column = worksheet.Columns(indices[0], indices[1]);
            
            Assert.AreEqual(indices[0], column.Min.Value);
            Assert.AreEqual(indices[1], column.Max.Value);
        }

        [TestMethod()]
        public void ColumnsTestOneIndex()
        {
            var worksheet = new Worksheet(new SheetData());

            Assert.AreSame(worksheet.Columns(1, 1), worksheet.Columns(1));
        }

        [TestMethod()]
        public void ColumnRangeTest()
        {
            var range = new Worksheet(new SheetData()).ColumnRange(2);

            Assert.AreEqual(2U, range.StartColumn);
            Assert.AreEqual(2U, range.EndColumn);
        }

        [TestMethod()]
        public void RowsTest()
        {
            var worksheet = new Worksheet(new SheetData());
            var row = worksheet.Rows(1);

            Assert.AreEqual(1U, row.RowIndex.Value);
            Assert.AreSame(row, worksheet.Rows(1));
        }

        [TestMethod()]
        public void RowRangeTest()
        {
            var range = new Worksheet(new SheetData()).RowRange(2);

            Assert.AreEqual(2U, range.StartRow);
            Assert.AreEqual(2U, range.EndRow);
        }

        [TestMethod()]
        public void CellsTest()
        {
            var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "D5" }) { RowIndex = 5 }));

            Assert.AreEqual("D5", worksheet.Cells(5, 4).CellReference.Value);
        }

        [TestMethod()]
        public void CellRangeTest()
        {
            var range = new Worksheet(new SheetData()).CellRange(1, 2, 6, 4);

            Assert.AreEqual(1U, range.StartRow);
            Assert.AreEqual(2U, range.StartColumn);
            Assert.AreEqual(6U, range.EndRow);
            Assert.AreEqual(5U, range.EndColumn);
        }

        [TestMethod()]
        public void CellRangeTestStartOnly()
        {
            var range = new Worksheet(new SheetData()).CellRange(6, 5);

            Assert.AreEqual(6U, range.StartRow);
            Assert.AreEqual(5U, range.StartColumn);
            Assert.AreEqual(6U, range.EndRow);
            Assert.AreEqual(5U, range.EndColumn);
        }

        [TestMethod()]
        public void NameTestGet()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet();
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;

                Assert.AreEqual("TestSheet", worksheet.Name());
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void NameTestSet()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet();
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;

                worksheet.Name("NewName");

                Assert.AreEqual("NewName", document.WorkbookPart.Workbook.Elements<Sheets>().Single().GetFirstChild<Sheet>().Name.Value);
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
        public void AddHyperlinkTest(bool applyHyperlinkStyle)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "ReferenceSheet" }, new Sheet() { Id = "rId2", SheetId = 2U, Name = "TargetSheet" }));
                var referenceWorksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }) { RowIndex = 1U }));
                var targetWorksheet = new Worksheet();
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = referenceWorksheet;
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId2").Worksheet = targetWorksheet;
                var stylesheet = new Stylesheet(
                    new Fonts(new Font()),
                    new Fills(new Fill()),
                    new Borders(new Border()),
                    new CellStyleFormats(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U }),
                    new CellFormats(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U }));
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;

                referenceWorksheet.AddHyperlink(1, 1, "TargetWorksheet", 2, 2, applyHyperlinkStyle);

                var hyperlink = referenceWorksheet.GetFirstChild<Hyperlinks>()?.GetFirstChild<Hyperlink>();
                var hyperlinkCell = referenceWorksheet.GetFirstChild<SheetData>().GetFirstChild<Row>().Elements<Cell>().Single(c => c.CellReference == "A1");

                Assert.IsNotNull(hyperlink);
                Assert.AreEqual(hyperlinkCell.CellReference.Value, hyperlink.Reference.Value);
                Assert.AreEqual("TargetWorksheet!B2", hyperlink.Location.Value);

                if (applyHyperlinkStyle)
                {
                    Assert.IsNotNull(hyperlinkCell.StyleIndex);
                    Assert.AreEqual(stylesheet.GetFirstChild<CellStyles>().Elements<CellStyle>().Single(s => s.BuiltinId == 8U).FormatId.Value, stylesheet.GetFirstChild<CellFormats>().Elements<CellFormat>().ElementAt((int)hyperlinkCell.StyleIndex.Value).FormatId.Value);
                }
                else
                {
                    Assert.IsNull(hyperlinkCell.StyleIndex);
                }
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void AddImageTest()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            Stream sourceStream = null;
            Stream resultStream = null;
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData(new Row(new Cell() { CellReference = "A1" }) { RowIndex = 1U }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                //sourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("SamuraiTools.OpenXmlTests.logo.jpg");
                sourceStream = new MemoryStream();
                System.Drawing.SystemIcons.WinLogo.Save(sourceStream);

                worksheet.AddImage(sourceStream, string.Empty, 1, 1);

                resultStream = worksheet.WorksheetPart.DrawingsPart?.ImageParts.ElementAt(0).GetStream();
                Assert.IsNotNull(resultStream);
                Assert.AreEqual(sourceStream.Length, resultStream.Length);

                var anchor = worksheet.WorksheetPart.DrawingsPart.WorksheetDrawing?.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor>().SingleOrDefault();
                Assert.IsNotNull(anchor);
                Assert.AreEqual("0", anchor.FromMarker.RowId.InnerText);
                Assert.AreEqual("0", anchor.FromMarker.ColumnId.InnerText);

                Assert.IsTrue(worksheet.Elements<Drawing>().Any());
            }
            finally
            {
                resultStream?.Dispose();
                sourceStream?.Dispose();
                document.Dispose();
                documentStream.Dispose();
            }
        }
    }

    [TestClass()]
    public class RowExtensionTest
    {
        [TestMethod()]
        public void CellsTest()
        {
            var worksheet = new Worksheet(new SheetData());
            var row = worksheet.GetFirstChild<SheetData>().AppendChild(new Row() { RowIndex = 1U });
            var cell = row.AppendChild(new Cell() { CellReference = "A1" });

            var cell1 = row.Cells(1);
            var cell2 = row.Cells(2);

            Assert.AreSame(cell, cell1);
            Assert.AreEqual("B1", cell2.CellReference.Value);
            Assert.AreEqual(2, row.ChildElements.Count);
        }

        [TestMethod()]
        public void ApplyCellStyleTest()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new SheetData());
                var row = worksheet.GetFirstChild<SheetData>().AppendChild(new Row(new Cell() { CellReference = "A1" }) { RowIndex = 1U });
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;

                row.ApplyCellStyle(BuiltinCellStyle.Hyperlink);

                var cellStyle = stylesheet.Elements<CellStyles>().Single().Elements<CellStyle>().SingleOrDefault(s => s.BuiltinId.Value == (uint)BuiltinCellStyle.Hyperlink);
                var rowCellFormat = stylesheet.Elements<CellFormats>().Single().Elements<CellFormat>().ElementAt((int)row.StyleIndex.Value);
                var cellFormat = stylesheet.Elements<CellFormats>().Single().Elements<CellFormat>().ElementAt((int)row.GetFirstChild<Cell>().StyleIndex.Value);

                Assert.AreEqual(cellStyle.FormatId.Value, rowCellFormat.FormatId.Value);
                Assert.AreEqual(cellStyle.FormatId.Value, cellFormat.FormatId.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }
    }

    [TestClass()]
    public class ColumnExtensionTest
    {
        [TestMethod()]
        public void AutoFitTest()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new Columns(), new SheetData(new Row(
                    new Cell() { CellReference = "A1", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("klbkbq98k28ghboiq7t54inurkeq")) },
                    new Cell() { CellReference = "A2", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("ft9huveahgyig8hgq39j8hwjniogaihu379")) })));
                var column = worksheet.GetFirstChild<Columns>().AppendChild(new Column() { Min = 1, Max = 1 });
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font() { FontName = new FontName() { Val = "Calibri" }, FontSize = new FontSize() { Val = 11D } }), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;

                column.AutoFit();

                double expectedWidth = 36.28515625;
                Assert.AreEqual(expectedWidth, column.Width.Value, expectedWidth * 0.15);
                Assert.IsTrue(column.CustomWidth.Value);
                Assert.IsTrue(column.BestFit.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void ApplyCellStyleTest()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                var worksheet = new Worksheet(new Columns(), new SheetData());
                var column = worksheet.GetFirstChild<Columns>().AppendChild(new Column() { Min = 1, Max = 2 });
                var row = worksheet.GetFirstChild<SheetData>().AppendChild(new Row(new Cell() { CellReference = "A1" }) { RowIndex = 1U });
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = worksheet;
                var stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(), new CellFormats());
                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = stylesheet;

                column.ApplyCellStyle(BuiltinCellStyle.Hyperlink);

                var cellStyle = stylesheet.Elements<CellStyles>().Single().Elements<CellStyle>().SingleOrDefault(s => s.BuiltinId.Value == (uint)BuiltinCellStyle.Hyperlink);
                var column1CellFormat = stylesheet.Elements<CellFormats>().Single().Elements<CellFormat>().ElementAt((int)worksheet.GetFirstChild<Columns>().Elements<Column>().Single(c => c.Min.Value == 1 && c.Max.Value == 1).Style.Value);
                var column2CellFormat = stylesheet.Elements<CellFormats>().Single().Elements<CellFormat>().ElementAt((int)worksheet.GetFirstChild<Columns>().Elements<Column>().Single(c => c.Min.Value == 2 && c.Max.Value == 2).Style.Value);
                var cellFormat = stylesheet.Elements<CellFormats>().Single().Elements<CellFormat>().ElementAt((int)row.GetFirstChild<Cell>().StyleIndex.Value);

                Assert.AreEqual(cellStyle.FormatId.Value, column1CellFormat.FormatId.Value);
                Assert.AreEqual(cellStyle.FormatId.Value, column2CellFormat.FormatId.Value);
                Assert.AreEqual(cellStyle.FormatId.Value, cellFormat.FormatId.Value);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }
    }
}