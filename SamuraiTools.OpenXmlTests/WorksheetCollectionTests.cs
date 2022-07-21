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
    public class WorksheetCollectionTests
    {
        [TestMethod()]
        [DataRow("NewSheet")]
        [DataRow("")]
        [DataRow(null)]
        public void AddTest(string name)
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = new Worksheet();
                var worksheets = new WorksheetCollection(document);
                var worksheet = new Worksheet();

                worksheets.Add(worksheet, name);
                
                Assert.AreSame(worksheet, document.WorkbookPart.WorksheetParts.Last().Worksheet);
                if (!string.IsNullOrEmpty(name))
                {
                    Assert.AreEqual(name, document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Last().Name.Value);
                }
                
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void ClearTest()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = new Worksheet();
                var worksheets = new WorksheetCollection(document);

                worksheets.Clear();

                Assert.IsFalse(document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Any());
                Assert.IsFalse(document.WorkbookPart.WorksheetParts.Any());
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void ContainsTestWorksheet()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = new Worksheet();
                var worksheets = new WorksheetCollection(document);

                var containedWorksheet = document.WorkbookPart.WorksheetParts.First().Worksheet;
                var nonContainedWorksheet = new Worksheet();

                Assert.IsTrue(worksheets.Contains(containedWorksheet));
                Assert.IsFalse(worksheets.Contains(nonContainedWorksheet));
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void ContainsTestName()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = new Worksheet();
                var worksheets = new WorksheetCollection(document);

                var containedWorksheet = "TestSheet";
                var nonContainedWorksheet = "NonContained";

                Assert.IsTrue(worksheets.Contains(containedWorksheet));
                Assert.IsFalse(worksheets.Contains(nonContainedWorksheet));
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void CopyToTest()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = new Worksheet();
                var worksheets = new WorksheetCollection(document);
                var worksheetArray = new Worksheet[2];
                worksheetArray[0] = new Worksheet();

                worksheets.CopyTo(worksheetArray, 1);

                Assert.AreSame(document.WorkbookPart.WorksheetParts.First().Worksheet, worksheetArray[1]);
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void RemoveTestWorksheet()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }, new Sheet() { Id = "rId2", SheetId = 2U, Name = "TestSheet2" }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = new Worksheet();
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId2").Worksheet = new Worksheet();
                var worksheets = new WorksheetCollection(document);
                var worksheet = document.WorkbookPart.WorksheetParts.First().Worksheet;

                worksheets.Remove(worksheet);

                Assert.AreEqual(1, document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Count());
                Assert.AreEqual(1, document.WorkbookPart.WorksheetParts.Count());
                Assert.IsNotNull(document.WorkbookPart.GetPartById("rId2"));
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void RemoveTestName()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }, new Sheet() { Id = "rId2", SheetId = 2U, Name = "TestSheet2" }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = new Worksheet();
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId2").Worksheet = new Worksheet();
                var worksheets = new WorksheetCollection(document);

                worksheets.Remove("TestSheet");

                Assert.AreEqual(1, document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Count());
                Assert.AreEqual(1, document.WorkbookPart.WorksheetParts.Count());
                Assert.IsNotNull(document.WorkbookPart.GetPartById("rId2"));
            }
            finally
            {
                document.Dispose();
                documentStream.Dispose();
            }
        }

        [TestMethod()]
        public void GetEnumeratorTest()
        {
            var documentStream = new MemoryStream();
            var document = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
            IEnumerator<Worksheet> enumerator = null;
            try
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets(new Sheet() { Id = "rId1", SheetId = 1U, Name = "TestSheet" }, new Sheet() { Id = "rId2", SheetId = 2U, Name = "TestSheet2" }));
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId1").Worksheet = new Worksheet();
                document.WorkbookPart.AddNewPart<WorksheetPart>("rId2").Worksheet = new Worksheet();
                var worksheets = new WorksheetCollection(document);

                enumerator = worksheets.GetEnumerator();
                enumerator.MoveNext();

                Assert.IsTrue(enumerator.MoveNext());
                Assert.IsFalse(enumerator.MoveNext());
            }
            finally
            {
                enumerator?.Dispose();
                document.Dispose();
                documentStream.Dispose();
            }
        }
    }
}