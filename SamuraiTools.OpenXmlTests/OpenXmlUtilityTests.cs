using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SamuraiTools.OpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace SamuraiTools.OpenXml.Tests
{
    [TestClass()]
    public class OpenXmlUtilityTests
    {
        [TestMethod()]
        public void GetNodeIndexTestFindExistingChild()
        {
            var parent = new Fonts();
            var child1 = parent.AppendChild(new Font() { Bold = new Bold(), FontName = new FontName() { Val = "Calibri" }, FontSize = new FontSize { Val = 11D } });

            var child2 = child1.Clone() as Font;
            child2.FontSize.Val = 14D;
            parent.Append(child2);
            
            Assert.AreEqual(0U, OpenXmlUtility.GetNodeIndex(child1, parent));
        }

        [TestMethod()]
        public void GetNodeIndexTestFindNewChild()
        {
            var parent = new Fonts();
            var child = new Font() { Bold = new Bold(), FontName = new FontName() { Val = "Calibri" }, FontSize = new FontSize { Val = 11D } };

            Assert.AreEqual(0U, OpenXmlUtility.GetNodeIndex(child, parent));
        }

        [TestMethod()]
        public void GetNodeIndexTestIncrementCount()
        {
            var parent = new Fonts();
            var child1 = new Font() { Bold = new Bold(), FontName = new FontName() { Val = "Calibri" }, FontSize = new FontSize { Val = 11D } };
            var child2 = child1.Clone() as Font;
            child2.FontSize.Val = 14D;

            OpenXmlUtility.GetNodeIndex(child1, parent);
            Assert.AreEqual(1U, parent.Count?.Value);

            OpenXmlUtility.GetNodeIndex(child2, parent);
            Assert.AreEqual(2U, parent.Count?.Value);
        }

        [TestMethod()]
        public void ConvertColorTestArgb()
        {
            int[] inputColorValues = { 255, 1, 128, 32 };
            string expectedValue = "FF018020";
            
            Assert.AreEqual(expectedValue, OpenXmlUtility.ConvertColor(inputColorValues[0], inputColorValues[1], inputColorValues[2], inputColorValues[3]));
        }

        [TestMethod()]
        public void ConvertColorTestRgb()
        {
            int[] inputColorValues = { 1, 128, 32 };
            string expectedValue = "FF018020";

            Assert.AreEqual(expectedValue, OpenXmlUtility.ConvertColor(inputColorValues[0], inputColorValues[1], inputColorValues[2]));
        }
    }
}