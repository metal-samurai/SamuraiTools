using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OXDrawing = DocumentFormat.OpenXml.Drawing;
using OXDrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace SamuraiTools.OpenXml
{
    public static class OpenXmlUtility
    {
        /// <summary>
        /// Get index of node within the specified parent node by comparing its XML to the current child nodes within parent. If the node isn't found it will be created and appended as the last child.
        /// </summary>
        /// <typeparam name="T">OpenXmlElement</typeparam>
        /// <typeparam name="TParent">OpenXmlElement</typeparam>
        /// <param name="node"></param>
        /// <param name="parentNode"></param>
        /// <returns>The index of the node within its parent.</returns>
        public static uint GetNodeIndex<T, TParent>(this T node, TParent parentNode) where T : OpenXmlElement where TParent : OpenXmlElement
        {
            OpenXmlElement newNode = parentNode.Elements<T>().FirstOrDefault(n => n.OuterXml == node.OuterXml);

            if (newNode == null)
            {
                newNode = parentNode.AppendChild(node);
                //some nodes have a Count property. if it's there i want to increment it.
                System.Reflection.PropertyInfo propertyInfo = parentNode.GetType().GetProperty("Count");
                if (propertyInfo != null)
                {
                    var currentValue = propertyInfo.GetValue(parentNode) as UInt32Value ?? new UInt32Value((uint)parentNode.ChildElements.Count - 1);
                    propertyInfo.SetValue(parentNode, new UInt32Value(currentValue.Value + 1U));
                }
            }

            return (uint)newNode.ElementsBefore().Count();
        }

        /// <summary>
        /// Convert separate ARGB integer values into a hex binary string used for formatting.
        /// </summary>
        /// <param name="alpha"></param>
        /// <param name="red"></param>
        /// <param name="green"></param>
        /// <param name="blue"></param>
        /// <returns></returns>
        public static string ConvertColor(int alpha, int red, int green, int blue)
        {
            return ((alpha << 24) + (red << 16) + (green << 8) + blue).ToString("X8");
        }

        /// <summary>
        /// Convert separate RGB integer values into a hex binary string used for formatting.
        /// </summary>
        /// <param name="red"></param>
        /// <param name="green"></param>
        /// <param name="blue"></param>
        /// <returns></returns>
        /// <remarks>Alpha will be set to 255.</remarks>
        public static string ConvertColor(int red, int green, int blue)
        {
            return ((255 << 24) + (red << 16) + (green << 8) + blue).ToString("X8");
        }
    }

    namespace Spreadsheet
    {
        public enum BuiltinCellStyle : uint
        {
            Normal = 0,
            Hyperlink = 8
        }

        public enum GradientFillType
        {
            Vertical,
            VerticalCenter,
            Horizontal,
            HorizontalCenter
        }

        public class AutoFitException : Exception
        {
            public string Worksheet { get; internal set; }
            public uint Column { get; internal set; }
            internal AutoFitException(string message) : base(message)
            { }
        }

        public static class SpreadsheetUtility
        {
            //credit to stackoverflow user graham
            /// <summary>
            /// Convert a column index number to its corresponding letter definition.
            /// </summary>
            /// <param name="columnNumber"></param>
            /// <returns></returns>
            public static string GetColumnLetter(uint columnNumber)
            {
                int dividend = (int)columnNumber;
                string columnName = string.Empty;
                int modulo;

                while (dividend > 0)
                {
                    modulo = (dividend - 1) % 26;
                    columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                    dividend = (int)((dividend - modulo) / 26);
                }

                return columnName;
            }

            /// <summary>
            /// Extract the column letter from a cell reference.
            /// </summary>
            /// <param name="cellReference"></param>
            /// <returns></returns>
            public static string GetColumnLetter(string cellReference)
            {
                //return new string(cellReference.Where(ch => !char.IsDigit(ch)).ToArray());
                return new System.Text.RegularExpressions.Regex("[A-Za-z]+").Match(cellReference).Value;
            }

            /// <summary>
            /// Extract the row index number from a cell reference.
            /// </summary>
            /// <param name="cellReference"></param>
            /// <returns></returns>
            public static uint GetRowIndex(string cellReference)
            {
                //return uint.Parse(new string(cellReference.Where(ch => char.IsDigit(ch)).ToArray()));
                return uint.Parse(new System.Text.RegularExpressions.Regex("\\d+").Match(cellReference).Value);
            }
        }

        public static class SpreadsheetDocumentExtension
        {
            /// <summary>
            /// Add a builtin cell style to a document's stylesheet.
            /// </summary>
            /// <param name="style">From the Open XML built-in cell style definitions</param>
            public static void AddCellStyle(this SpreadsheetDocument document, BuiltinCellStyle style)
            {
                Stylesheet stylesheet = document.WorkbookPart.WorkbookStylesPart.Stylesheet;
                //NumberingFormat numberingFormat;
                Font font;
                //Fill fill;
                //Border border;

                CellStyles cellStyles = stylesheet.Elements<CellStyles>().SingleOrDefault() ?? stylesheet.Elements<CellFormats>().Single().InsertAfterSelf(new CellStyles() { Count = 0U });
                if (cellStyles.Count == null) cellStyles.Count = (uint)cellStyles.ChildElements.Count;

                CellFormat cellStyleFormat;
                CellStyle cellStyle;

                //check if it already exists
                if (cellStyles.Elements<CellStyle>().SingleOrDefault(s => s.BuiltinId.Value == (uint)style) != null)
                {
                    return;
                }

                switch (style)
                {
                    case BuiltinCellStyle.Normal:
                        cellStyleFormat = new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U };

                        cellStyle = new CellStyle() { Name = "Normal", FormatId = cellStyleFormat.GetNodeIndex(stylesheet.Elements<CellStyleFormats>().Single()), BuiltinId = (uint)style };
                        cellStyles.AppendChild(cellStyle);
                        cellStyles.Count++;
                        break;

                    case BuiltinCellStyle.Hyperlink:
                        /*if (document.WorkbookPart.ThemePart == null)
                        {
                            document.WorkbookPart.AddNewPart<ThemePart>();
                        }

                        document.WorkbookPart.ThemePart.Theme ??= new OXDrawing::Theme() { Name = "SamuraiTools" };
                        document.WorkbookPart.ThemePart.Theme.ThemeElements ??= new OXDrawing::ThemeElements();
                        document.WorkbookPart.ThemePart.Theme.ThemeElements.ColorScheme ??= new OXDrawing::ColorScheme() { Name = "SamuraiTools" };
                        document.WorkbookPart.ThemePart.Theme.ThemeElements.ColorScheme.Hyperlink ??= new OXDrawing::Hyperlink(new OXDrawing::RgbColorModelHex() { Val = "0563C1" });
                        */
                        font = document.CreateStyleElementTemplate<Font>(null);
                        if (document.WorkbookPart.ThemePart?.Theme?.ThemeElements?.ColorScheme?.Hyperlink != null)
                        {
                            font.Color = new Color() { Theme = (uint)document.WorkbookPart.ThemePart.Theme.ThemeElements.ColorScheme.Hyperlink.ElementsBefore().Count() };
                        }
                        else
                        {
                            font.Color = new Color() { Rgb = "FF0563C1" };
                        }
                        
                        font.Underline = new Underline();

                        cellStyleFormat = new CellFormat() { NumberFormatId = 0U, FontId = font.GetNodeIndex(stylesheet.Elements<Fonts>().Single()), FillId = 0U, BorderId = 0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };

                        cellStyle = new CellStyle() { Name = "Hyperlink", FormatId = cellStyleFormat.GetNodeIndex(stylesheet.Elements<CellStyleFormats>().Single()), BuiltinId = (uint)style };
                        cellStyles.AppendChild(cellStyle);
                        cellStyles.Count++;
                        break;
                }
            }

            /// <summary>
            /// Apply builtin cell style to an existing CellFormat.
            /// </summary>
            /// <param name="document"></param>
            /// <param name="cellFormat"></param>
            /// <param name="style"></param>
            public static void ApplyCellStyle(this SpreadsheetDocument document, CellFormat cellFormat, BuiltinCellStyle style)
            {
                Stylesheet stylesheet = document.WorkbookPart.WorkbookStylesPart.Stylesheet;
                CellStyle cellStyle = null;
                System.Action SetCellStyle = delegate ()
                {
                    cellStyle = stylesheet.Elements<CellStyles>().SingleOrDefault()?.Elements<CellStyle>().SingleOrDefault(s => s.BuiltinId == (uint)style);
                };

                SetCellStyle();
                if (cellStyle == null)
                {
                    document.AddCellStyle(style);
                    SetCellStyle();
                }

                if (cellStyle == null)
                {
                    return;
                }

                CellFormat cellStyleFormat = stylesheet.Elements<CellStyleFormats>().Single().Elements<CellFormat>().ElementAt((int)cellStyle.FormatId.Value);

                if (cellStyleFormat.ApplyNumberFormat?.Value ?? true)
                {
                    cellFormat.NumberFormatId = cellStyleFormat.NumberFormatId;
                    cellFormat.ApplyNumberFormat = null;
                }

                if (cellStyleFormat.ApplyFont?.Value ?? true)
                {
                    cellFormat.FontId = cellStyleFormat.FontId;
                    cellFormat.ApplyFont = null;
                }

                if (cellStyleFormat.ApplyFill?.Value ?? true)
                {
                    cellFormat.FillId = cellStyleFormat.FillId;
                    cellFormat.ApplyFill = null;
                }

                if (cellStyleFormat.ApplyBorder?.Value ?? true)
                {
                    cellFormat.BorderId = cellStyleFormat.BorderId;
                    cellFormat.ApplyBorder = null;
                }

                if (cellStyleFormat.ApplyAlignment?.Value ?? true)
                {
                    cellFormat.Alignment = cellStyleFormat.Alignment;
                    cellFormat.ApplyAlignment = null;
                }

                if (cellStyleFormat.ApplyProtection?.Value ?? true)
                {
                    cellFormat.Protection = cellStyleFormat.Protection;
                    cellFormat.ApplyProtection = null;
                }

                cellFormat.FormatId = cellStyle.FormatId;
            }

            /// <summary>
            /// Create and add a WorkbookStylesPart and Stylesheet with the minimum components for CellFormats to function, including a starting Font object for Calibri 11.
            /// </summary>
            public static void AddNewStylesheet(this SpreadsheetDocument document)
            {
                document.AddNewStylesheet(new Font() { FontName = new FontName() { Val = "Calibri" }, FontSize = new FontSize { Val = 11D } });
            }

            /// <summary>
            /// Create and add a WorkbookStylesPart and Stylesheet with the minimum components for CellFormats to function.
            /// </summary>
            /// <param name="document"></param>
            /// <param name="elements">Additional elements of type Font, Fill, or Border to include in the Stylesheet.</param>
            public static void AddNewStylesheet(this SpreadsheetDocument document, params OpenXmlElement[] elements)
            {
                var fontElements = elements.OfType<Font>();
                var fillElements = elements.OfType<Fill>();
                var borderElements = elements.OfType<Border>();

                //a Font should always indicate name and size, so only create a blank one if necessary.
                Fonts fonts = new Fonts(fontElements);
                if (fonts.ChildElements.Count == 0)
                {
                    fonts.AppendChild(new Font());
                }
                fonts.Count = (uint)fonts.ChildElements.Count;
                
                //the first 2 Fill objects are expected to be a certain way.
                Fills fills = new Fills(new Fill(new PatternFill() { PatternType = new EnumValue<PatternValues>(PatternValues.None) }), new Fill(new PatternFill() { PatternType = new EnumValue<PatternValues>(PatternValues.Gray125) }));
                fills.Append(fillElements.Where(x => !fills.ChildElements.Any(y => y.OuterXml == x.OuterXml)));
                fills.Count = (uint)fills.ChildElements.Count;

                //start with an empty border and append any additional ones.
                Borders borders = new Borders(new Border());
                borders.Append(borderElements.Where(x => !borders.ChildElements.Any(y => y.OuterXml == x.OuterXml)));
                borders.Count = (uint)borders.ChildElements.Count;

                document.WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(
                    fonts,
                    fills,
                    borders,
                    new CellStyleFormats(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U }) { Count = 1U },
                    new CellFormats(new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U }) { Count = 1U });

                document.AddCellStyle(BuiltinCellStyle.Normal);
            }

            /// <summary>
            /// Create a new style element based on the given StyleIndex.
            /// If that index is null, create a new basic element appropriate for the type, using the document's Normal style.
            /// If the Normal style is not defined, use the first element of the appropriate type on the Stylesheet.
            /// If there are no elements of that type defined, return a new blank instance of that element type.
            /// </summary>
            /// <typeparam name="T">OpenXmlElement</typeparam>
            /// <param name="document"></param>
            /// <param name="styleIndex">The index pointing to a CellFormat to use as a template if not null.</param>
            /// <returns></returns>
            public static T CreateStyleElementTemplate<T>(this SpreadsheetDocument document, UInt32Value styleIndex) where T : OpenXmlElement
            {
                Stylesheet stylesheet = document.WorkbookPart.WorkbookStylesPart.Stylesheet;

                System.Func<T> GetExistingNode = delegate ()
                {
                    CellFormat currentFormat = stylesheet.Elements<CellFormats>().Single().Elements<CellFormat>().ElementAt((int)styleIndex.Value);

                    if (typeof(T) == typeof(CellFormat))
                    {
                        return currentFormat.CloneNode(true) as T;
                    }
                    else if (typeof(T) == typeof(Font))
                    {
                        return stylesheet.Elements<Fonts>().Single().Elements<Font>().ElementAt((int)currentFormat.FontId.Value).CloneNode(true) as T;
                    }
                    else if (typeof(T) == typeof(Fill))
                    {
                        return stylesheet.Elements<Fills>().Single().Elements<Fill>().ElementAt((int)currentFormat.FillId.Value).CloneNode(true) as T;
                    }
                    else if (typeof(T) == typeof(Border))
                    {
                        return stylesheet.Elements<Borders>().Single().Elements<Border>().ElementAt((int)currentFormat.BorderId.Value).CloneNode(true) as T;
                    }
                    else
                    {
                        return null;
                    }
                };

                System.Func<T> GetNewNode = delegate ()
                {
                    CellStyle normalStyle = document.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<CellStyles>().FirstOrDefault()?.Elements<CellStyle>().SingleOrDefault(s => s.BuiltinId == (uint)BuiltinCellStyle.Normal);
                    CellFormat normalFormat = normalStyle == null ? null : document.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<CellStyleFormats>().Single().Elements<CellFormat>().ElementAt((int)normalStyle.FormatId.Value);

                    if (typeof(T) == typeof(CellFormat))
                    {
                        if (normalFormat != null)
                        {
                            return new CellFormat() { NumberFormatId = normalFormat.NumberFormatId, FontId = normalFormat.FontId, FillId = normalFormat.FillId, BorderId = normalFormat.BorderId, FormatId = normalStyle.FormatId } as T;
                        }
                        else
                        {
                            return new CellFormat() { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 } as T;
                        }
                    }
                    else if (typeof(T) == typeof(Font))
                    {
                        if (normalFormat != null)
                        {
                            return document.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<Fonts>().Single().Elements<Font>().ElementAt((int)normalFormat.FontId.Value).CloneNode(true) as T;
                        }
                        else
                        {
                            return (document.WorkbookPart.WorkbookStylesPart?.Stylesheet.Elements<Fonts>().SingleOrDefault()?.Elements<Font>().FirstOrDefault() ?? new Font()) as T;
                        }
                    }
                    else if (typeof(T) == typeof(Fill))
                    {
                        if (normalFormat != null)
                        {
                            return document.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<Fills>().Single().Elements<Fill>().ElementAt((int)normalFormat.FillId.Value).CloneNode(true) as T;
                        }
                        else
                        {
                            return (document.WorkbookPart.WorkbookStylesPart?.Stylesheet.Elements<Fills>().SingleOrDefault()?.Elements<Fill>().FirstOrDefault() ?? new Fill()) as T;
                        }
                    }
                    else if (typeof(T) == typeof(Border))
                    {
                        if (normalFormat != null)
                        {
                            return document.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<Borders>().Single().Elements<Border>().ElementAt((int)normalFormat.BorderId.Value).CloneNode(true) as T;
                        }
                        else
                        {
                            return (document.WorkbookPart.WorkbookStylesPart?.Stylesheet.Elements<Borders>().SingleOrDefault()?.Elements<Border>().FirstOrDefault() ?? new Border()) as T;
                        }
                    }
                    else
                    {
                        return null;
                    }
                };

                if (styleIndex?.HasValue ?? false)
                {
                    return GetExistingNode();
                }
                else
                {
                    return GetNewNode();
                }
            }
            
            /// <summary>
            /// Add the minimum required components for a new document to be useable. A document must have a Workbook with at least one Worksheet.
            /// </summary>
            /// <param name="document">A blank document.</param>
            public static void PrepNewDocument(this SpreadsheetDocument document)
            {
                document.AddWorkbookPart().Workbook = new Workbook(new Sheets());
                document.Worksheets().AddNew(null);
            }

            //from MSDN
            /// <summary>
            /// Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
            /// reference the specified SharedStringItem and removes the item.
            /// </summary>
            /// <param name="shareStringId">ID of shared string to check</param>
            public static void RemoveSharedStringItem(this SpreadsheetDocument document, uint shareStringId)
            {
                bool remove = true;

                foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
                {
                    Worksheet worksheet = part.Worksheet;
                    foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
                    {
                        // Verify if other cells in the document reference the item.
                        if (cell.DataType != null &&
                            cell.DataType.Value == CellValues.SharedString &&
                            cell.CellValue.Text == shareStringId.ToString())
                        {
                            // Other cells in the document still reference the item. Do not remove the item.
                            remove = false;
                            break;
                        }
                    }

                    if (!remove)
                    {
                        break;
                    }
                }

                // Other cells in the document do not reference the item. Remove the item.
                if (remove)
                {
                    SharedStringTablePart shareStringTablePart = document.WorkbookPart.SharedStringTablePart;
                    if (shareStringTablePart == null)
                    {
                        return;
                    }

                    SharedStringItem item = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt((int)shareStringId);
                    if (item != null)
                    {
                        item.Remove();
                        if (shareStringTablePart.SharedStringTable.Count == null) shareStringTablePart.SharedStringTable.Count = (uint)shareStringTablePart.SharedStringTable.ChildElements.Count;
                        shareStringTablePart.SharedStringTable.Count--;

                        // Refresh all the shared string references.
                        foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
                        {
                            Worksheet worksheet = part.Worksheet;
                            foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
                            {
                                if (cell.DataType != null &&
                                    cell.DataType.Value == CellValues.SharedString)
                                {
                                    int itemIndex = int.Parse(cell.CellValue.Text);
                                    if (itemIndex > shareStringId)
                                    {
                                        cell.CellValue.Text = (itemIndex - 1).ToString();
                                    }
                                }
                            }
                            //worksheet.Save();
                        }

                        //document.WorkbookPart.SharedStringTablePart.SharedStringTable.Save();
                    }
                }
            }

            /// <summary>
            /// Get the collection of Worksheet objects for this document.
            /// </summary>
            /// <param name="document"></param>
            /// <returns>A new WorksheetCollection object.</returns>
            public static WorksheetCollection Worksheets(this SpreadsheetDocument document)
            {
                return new WorksheetCollection(document);
            }

            /// <summary>
            /// Get the Worksheet with the given name.
            /// </summary>
            /// <param name="document"></param>
            /// <param name="name"></param>
            /// <returns></returns>
            public static Worksheet Worksheets(this SpreadsheetDocument document, string name)
            {
                return new WorksheetCollection(document)[name];
            }

            /// <summary>
            /// Get the worksheet at the given index.
            /// </summary>
            /// <param name="document"></param>
            /// <param name="index"></param>
            /// <returns></returns>
            public static Worksheet Worksheets(this SpreadsheetDocument document, int index)
            {
                return new WorksheetCollection(document)[index];
            }
        }

        public static class WorksheetExtension
        {
            /// <summary>
            /// Get the Columns element of the worksheet.
            /// </summary>
            /// <param name="worksheet"></param>
            /// <returns></returns>
            public static Columns Columns(this Worksheet worksheet)
            {
                return worksheet.Elements<Columns>().SingleOrDefault() ?? worksheet.InsertBefore(new Columns(), worksheet.Elements<SheetData>().Single());
            }

            /// <summary>
            /// Get the Column element with the given minimum and maximum index.
            /// </summary>
            /// <param name="worksheet"></param>
            /// <param name="min"></param>
            /// <param name="max"></param>
            /// <returns></returns>
            public static Column Columns(this Worksheet worksheet, uint min, uint max)
            {
                Columns columns = worksheet.Columns();
                //if there's no width specified it's treated as 0 apparently.
                return columns.Elements<Column>().Where(c => c.Min == min && c.Max == max).FirstOrDefault() ?? columns.AppendChild(new Column() { Min = min, Max = max, Width = 8D, CustomWidth = false });
            }

            /// <summary>
            /// Get the Column element with the given index as minimum and maximum index.
            /// </summary>
            /// <param name="worksheet"></param>
            /// <param name="index"></param>
            /// <returns></returns>
            public static Column Columns(this Worksheet worksheet, uint index)
            {
                return worksheet.Columns(index, index);
            }

            /// <summary>
            /// Get a Range representing the column at the given index.
            /// </summary>
            /// <param name="worksheet"></param>
            /// <param name="index"></param>
            /// <returns></returns>
            public static Range ColumnRange(this Worksheet worksheet, uint index)
            {
                return new Range(worksheet, null, index, null, 1);
            }

            /// <summary>
            /// Find Row at the given index on this Worksheet.
            /// </summary>
            /// <param name="index"></param>
            /// <returns></returns>
            public static Row Rows(this Worksheet worksheet, uint index)
            {
                //rows are written in order by index number
                Row row = null;
                Row refRow = null;
                SheetData sheetData = worksheet.Elements<SheetData>().Single();

                refRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex >= index);
                if (refRow == null || refRow.RowIndex > index)
                {
                    row = sheetData.InsertBefore(new Row() { RowIndex = index }, refRow);
                }
                else
                {
                    row = refRow;
                }

                return row;
            }

            /// <summary>
            /// Get a Range representing the row at the given index.
            /// </summary>
            /// <param name="worksheet"></param>
            /// <param name="index"></param>
            /// <returns></returns>
            public static Range RowRange(this Worksheet worksheet, uint index)
            {
                return new Range(worksheet, index, null, 1, null);
            }

            /// <summary>
            /// Get the Cell at the given row and column index.
            /// </summary>
            /// <param name="worksheet"></param>
            /// <param name="rowIndex"></param>
            /// <param name="columnIndex"></param>
            /// <returns></returns>
            public static Cell Cells(this Worksheet worksheet, uint rowIndex, uint columnIndex)
            {
                return worksheet.Rows(rowIndex).Cells(columnIndex);
            }

            /// <summary>
            /// Get a Range for all cells on this Worksheet between the starting point and the number of rows and columns.
            /// </summary>
            /// <param name="worksheet"></param>
            /// <param name="startRow"></param>
            /// <param name="startColumn"></param>
            /// <param name="rowCount"></param>
            /// <param name="columnCount"></param>
            /// <returns>A new Range object.</returns>
            public static Range CellRange(this Worksheet worksheet, uint startRow, uint startColumn, uint rowCount, uint columnCount)
            {
                return new Range(worksheet, startRow, startColumn, rowCount, columnCount);
            }

            /// <summary>
            /// Get a Range for the cell on this Worksheet at the given row and column index.
            /// </summary>
            /// <param name="worksheet"></param>
            /// <param name="startRow"></param>
            /// <param name="startColumn"></param>
            /// <returns>A new Range object.</returns>
            public static Range CellRange(this Worksheet worksheet, uint startRow, uint startColumn)
            {
                return new Range(worksheet, startRow, startColumn);
            }

            /// <summary>
            /// Get the name of this Worksheet.
            /// </summary>
            /// <param name="worksheet"></param>
            /// <returns>The display name for the Worksheet when the document is opened.</returns>
            public static string Name(this Worksheet worksheet)
            {
                WorkbookPart workbookPart = ((SpreadsheetDocument)worksheet.WorksheetPart.OpenXmlPackage).WorkbookPart;

                return workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Single(s => s.Id == workbookPart.GetIdOfPart(worksheet.WorksheetPart)).Name;
            }

            /// <summary>
            /// Change the name for this Worksheet.
            /// </summary>
            /// <param name="worksheet"></param>
            /// <param name="newName"></param>
            public static void Name(this Worksheet worksheet, string newName)
            {
                WorkbookPart workbookPart = ((SpreadsheetDocument)worksheet.WorksheetPart.OpenXmlPackage).WorkbookPart;

                workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Single(s => s.Id == workbookPart.GetIdOfPart(worksheet.WorksheetPart)).Name = newName;
            }

            internal static void AddHyperlink(this Worksheet worksheet, string reference, string target)
            {
                (worksheet.Elements<Hyperlinks>().FirstOrDefault() ?? worksheet.Elements<SheetData>().Single().InsertAfterSelf(new Hyperlinks())).AppendChild(new Hyperlink() { Reference = reference, Location = target });
            }

            /// <summary>
            /// Add a Hyperlink to the reference cell which points to the target cell.
            /// </summary>
            /// <param name="worksheet"></param>
            /// <param name="referenceRow"></param>
            /// <param name="referenceColumn"></param>
            /// <param name="targetWorksheetName"></param>
            /// <param name="targetRow"></param>
            /// <param name="targetColumn"></param>
            /// <param name="formatCell">Whether to apply hyperlink styling to the reference cell.</param>
            public static void AddHyperlink(this Worksheet worksheet, uint referenceRow, uint referenceColumn, string targetWorksheetName, uint targetRow, uint targetColumn, bool formatCell)
            {
                worksheet.AddHyperlink(SpreadsheetUtility.GetColumnLetter(referenceColumn) + referenceRow.ToString(), (string.IsNullOrEmpty(targetWorksheetName) ? string.Empty : targetWorksheetName + "!") + SpreadsheetUtility.GetColumnLetter(targetColumn) + targetRow.ToString());
                if (formatCell)
                {
                    worksheet.Cells(referenceRow, referenceColumn).ApplyCellStyle(BuiltinCellStyle.Hyperlink);
                }
            }

            //credit to stackoverflow user benjamin krupp for the general approach
            /// <summary>
            /// Add and image to this Worksheet.
            /// </summary>
            /// <param name="worksheet"></param>
            /// <param name="imageStream">Stream containing image data.</param>
            /// <param name="imgDesc">Description of the image.</param>
            /// <param name="rowIndex">Row of the cell at which to place the image.</param>
            /// <param name="colIndex">Column of the cell at which to place the image.</param>
            public static void AddImage(this Worksheet worksheet, Stream imageStream, string imgDesc, uint rowIndex, uint colIndex)
            {
                WorksheetPart worksheetPart = worksheet.WorksheetPart;
                DrawingsPart drawingsPart = worksheetPart.DrawingsPart;
                ImagePart imagePart;

                long extentsCx;
                long extentsCy;

                ImagePartType imagePartType;

                using (System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(imageStream))
                {
                    if (System.Drawing.Imaging.ImageFormat.Bmp.Equals(bitmap.RawFormat))
                        imagePartType = ImagePartType.Bmp;
                    else if (System.Drawing.Imaging.ImageFormat.Gif.Equals(bitmap.RawFormat))
                        imagePartType = ImagePartType.Gif;
                    else if (System.Drawing.Imaging.ImageFormat.Png.Equals(bitmap.RawFormat))
                        imagePartType = ImagePartType.Png;
                    else if (System.Drawing.Imaging.ImageFormat.Tiff.Equals(bitmap.RawFormat))
                        imagePartType = ImagePartType.Tiff;
                    else if (System.Drawing.Imaging.ImageFormat.Icon.Equals(bitmap.RawFormat))
                        imagePartType = ImagePartType.Icon;
                    else if (System.Drawing.Imaging.ImageFormat.Jpeg.Equals(bitmap.RawFormat))
                        imagePartType = ImagePartType.Jpeg;
                    else if (System.Drawing.Imaging.ImageFormat.Emf.Equals(bitmap.RawFormat))
                        imagePartType = ImagePartType.Emf;
                    else if (System.Drawing.Imaging.ImageFormat.Wmf.Equals(bitmap.RawFormat))
                        imagePartType = ImagePartType.Wmf;
                    else
                        throw new Exception("Image type could not be determined.");

                    //calcuate canvas size in EMUs, 914400 per inch.
                    extentsCx = bitmap.Width * (long)(914400 / bitmap.HorizontalResolution);
                    extentsCy = bitmap.Height * (long)(914400 / bitmap.VerticalResolution);
                }

                if (drawingsPart == null)
                    drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

                if (!worksheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
                    worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });

                if (drawingsPart.WorksheetDrawing == null)
                    drawingsPart.WorksheetDrawing = new OXDrawingSpreadsheet::WorksheetDrawing();

                imagePart = drawingsPart.AddImagePart(imagePartType);
                imageStream.Position = 0;
                imagePart.FeedData(imageStream);

                IEnumerable<OXDrawing::NonVisualDrawingProperties> nvps = drawingsPart.WorksheetDrawing.Descendants<OXDrawing::NonVisualDrawingProperties>();
                uint nvpId = nvps.Count() > 0 ? (UInt32Value)nvps.Max(p => p.Id.Value) + 1 : 1U;

                drawingsPart.WorksheetDrawing.Append(new OXDrawingSpreadsheet::OneCellAnchor(
                    new OXDrawingSpreadsheet::FromMarker
                    {
                        RowId = new OXDrawingSpreadsheet::RowId((rowIndex - 1).ToString()),
                        ColumnId = new OXDrawingSpreadsheet::ColumnId((colIndex - 1).ToString()),
                        RowOffset = new OXDrawingSpreadsheet::RowOffset("0"),
                        ColumnOffset = new OXDrawingSpreadsheet::ColumnOffset("0")
                    },
                    new OXDrawingSpreadsheet::Extent { Cx = extentsCx, Cy = extentsCy },
                    new OXDrawingSpreadsheet::Picture(
                        new OXDrawingSpreadsheet::NonVisualPictureProperties(
                            new OXDrawingSpreadsheet::NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId, Description = imgDesc },
                            new OXDrawingSpreadsheet::NonVisualPictureDrawingProperties(new OXDrawing::PictureLocks { NoChangeAspect = true })
                        ),
                        new OXDrawingSpreadsheet::BlipFill(
                            new OXDrawing::Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = OXDrawing::BlipCompressionValues.Print },
                            new OXDrawing::Stretch(new OXDrawing::FillRectangle())
                        ),
                        new OXDrawingSpreadsheet::ShapeProperties(
                            new OXDrawing::Transform2D(
                                new OXDrawing::Offset { X = 0, Y = 0 },
                                new OXDrawing::Extents { Cx = extentsCx, Cy = extentsCy }
                            ),
                            new OXDrawing::PresetGeometry { Preset = OXDrawing::ShapeTypeValues.Rectangle }
                        )
                    ),
                    new OXDrawingSpreadsheet::ClientData()
                ));
            }

            /// <summary>
            /// Add and image to this Worksheet.
            /// </summary>
            /// <param name="worksheet"></param>
            /// <param name="imagePath">Full path to the image file.</param>
            /// <param name="imgDesc">Description of the image.</param>
            /// <param name="rowIndex">Row of the cell at which to place the image.</param>
            /// <param name="colIndex">Column of the cell at which to place the image.</param>
            public static void AddImage(this Worksheet worksheet, string imagePath, string imgDesc, uint rowIndex, uint colIndex)
            {
                using (var imageStream = new FileStream(imagePath, FileMode.Open))
                {
                    AddImage(worksheet, imageStream, imgDesc, rowIndex, colIndex);
                }
            }
        }

        public static class RowExtension
        {
            /// <summary>
            /// Get the Cell in this Row at the given column index.
            /// </summary>
            /// <param name="row"></param>
            /// <param name="columnIndex"></param>
            /// <returns></returns>
            public static Cell Cells(this Row row, uint columnIndex)
            {
                string cellReference = SpreadsheetUtility.GetColumnLetter(columnIndex) + row.RowIndex.ToString();

                // If there is not a cell with the specified column name, insert one.  
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell cell = null;
                Cell refCell = null;
                foreach (Cell c in row.Elements<Cell>())
                {
                    if (c.CellReference.Value == cellReference)
                    {
                        cell = c;
                        break;
                    }
                    else if (c.CellReference.Value.Length > cellReference.Length || (c.CellReference.Value.Length == cellReference.Length && string.Compare(c.CellReference.Value, cellReference, true) > 0))
                    {
                        refCell = c;
                        break;
                    }
                }
                if (cell == null)
                {
                    cell = row.InsertBefore(new Cell() { CellReference = cellReference }, refCell);
                    Column refColumn = row.Ancestors<Worksheet>().Single().Elements<Columns>().SingleOrDefault()?.Elements<Column>().FirstOrDefault(c => c.Min <= columnIndex && c.Max >= columnIndex && c.Style != null);
                    if (refColumn != null)
                    {
                        cell.StyleIndex = refColumn.Style;
                    }
                    else
                    {
                        cell.StyleIndex = row.StyleIndex;
                    }
                }

                return cell;
            }

            /// <summary>
            /// Apply a builtin cell style to this Row and its cells. The necessary CellStyle will be created if it doesn't yet exist in the StyleSheet.
            /// </summary>
            /// <param name="row"></param>
            /// <param name="style"></param>
            public static void ApplyCellStyle(this Row row, BuiltinCellStyle style)
            {
                new Range(row.Ancestors<Worksheet>().Single(), row.RowIndex, null).ApplyCellStyle(style);
            }
        }

        public static class ColumnExtension
        {
            /// <summary>
            /// Set the width of a column to display the longest text contained in any of its cells.
            /// </summary>
            /// <param name="column"></param>
            public static void AutoFit(this Column column)
            {
                Worksheet worksheet = column.Ancestors<Worksheet>().Single();
                SpreadsheetDocument document = (SpreadsheetDocument)worksheet.WorksheetPart.OpenXmlPackage;
                
                //Column width is a multiple of the width of a digit in the default font so we need to identify it.
                Font defaultFont = null;
                
                //First use Normal cell style if it's defined and useable.
                CellStyle normalStyle = document.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<CellStyles>().FirstOrDefault()?.Elements<CellStyle>().SingleOrDefault(s => s.BuiltinId == (uint)BuiltinCellStyle.Normal);
                if (normalStyle != null && (defaultFont = document.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<Fonts>().Single().Elements<Font>().ElementAt((int)document.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<CellStyleFormats>().Single().Elements<CellFormat>().ElementAt((int)normalStyle.FormatId.Value).FontId.Value)).FontName?.Val != null && defaultFont.FontSize?.Val != null)
                {
                }
                //Next try the first font available.
                else if ((defaultFont = document.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<Fonts>().Single().GetFirstChild<Font>()).FontName?.Val != null && defaultFont.FontSize?.Val != null)
                {
                }
                else
                {
                    throw new AutoFitException("Unable to determine font for AutoSize") { Worksheet = worksheet.Name(), Column = column.Min };
                }

                //Minimum and maxium character width in the default font (width of 'i' and '0').
                double minCharWidth = 0D;
                double maxCharWidth = 0D;

                //Excel renders each character with a whole number of pixels, but the number used seems chosen arbitrarily rather than using conventional rounding.
                //My best approximation is random rounding with a minimum.
                Random random = new Random();
                Func<double, double> RandomRound = delegate (double x)
                {
                    return Math.Max(random.NextDouble() > 0.5 ? Math.Ceiling(x) : Math.Truncate(x), minCharWidth);
                };

                System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(1, 1);
                System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(bitmap);

                try
                {
                    graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                    graphics.PageUnit = System.Drawing.GraphicsUnit.Pixel;

                    using (var drawFont = new System.Drawing.Font(defaultFont.FontName.Val.Value, (float)defaultFont.FontSize.Val.Value, System.Drawing.GraphicsUnit.Point))
                    {
                        minCharWidth = Math.Ceiling(graphics.MeasureString("i", drawFont, new System.Drawing.PointF(0F, 0F), System.Drawing.StringFormat.GenericTypographic).Width);
                        maxCharWidth = Math.Truncate(graphics.MeasureString("0", drawFont, new System.Drawing.PointF(0F, 0F), System.Drawing.StringFormat.GenericTypographic).Width);
                    }

                    for (uint i = column.Min; i <= column.Max; i++)
                    {
                        Column individualColumn = worksheet.Columns(i);
                        string columnLetter = SpreadsheetUtility.GetColumnLetter(i);
                        //Maximum width of any cell in this column.
                        double maxCellWidth = 0D;
                        //Text displayed in the cell.
                        string cellValue = string.Empty;
                        //Font used in an individual cell.
                        Font referenceFont = null;

                        foreach (var cell in worksheet.Descendants<Cell>().Where(c => SpreadsheetUtility.GetColumnLetter(c.CellReference) == columnLetter && c.InnerText.Length > 0))
                        {
                            cellValue = cell.InnerText;

                            if (cell.DataType != null)
                            {
                                switch (cell.DataType.Value)
                                {
                                    case CellValues.SharedString:
                                        cellValue = document.WorkbookPart.SharedStringTablePart.SharedStringTable.ElementAt(int.Parse(cell.InnerText)).InnerText;
                                        break;

                                    case CellValues.Boolean:
                                        cellValue = cell.InnerText == "0" ? "FALSE" : "TRUE";
                                        break;
                                }
                            }

                            referenceFont = cell.StyleIndex?.HasValue ?? false ? document.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<Fonts>().Single().Elements<Font>().ElementAt((int)document.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>().Single().Elements<CellFormat>().ElementAt((int)cell.StyleIndex.Value).FontId.Value) : defaultFont;

                            if (referenceFont?.FontName?.Val == null || referenceFont?.FontSize?.Val == null)
                            {
                                throw new AutoFitException("Unable to determine font for AutoSize") { Worksheet = worksheet.Name(), Column = i };
                            }

                            double textWidth = 0D;

                            using (var drawFont = new System.Drawing.Font(referenceFont.FontName.Val.Value, (float)referenceFont.FontSize.Val.Value, System.Drawing.GraphicsUnit.Point))
                            {
                                foreach (char c in cellValue.ToCharArray())
                                {
                                    textWidth += RandomRound(graphics.MeasureString(c.ToString(), drawFont, new System.Drawing.PointF(0F, 0F), System.Drawing.StringFormat.GenericTypographic).Width);
                                }
                            }
                            
                            maxCellWidth = Math.Max(maxCellWidth, textWidth);
                        }

                        if (maxCellWidth > 0)
                        {
                            //Excel uses a small amount (5px) of padding in cells.
                            individualColumn.Width = Math.Truncate((maxCellWidth + 5) / maxCharWidth * 256) / 256;
                            individualColumn.CustomWidth = true;
                            individualColumn.BestFit = true;
                        }
                    }
                }
                catch
                {
                    throw;
                }
                finally
                {
                    graphics.Dispose();
                    bitmap.Dispose();
                }
            }

            /// <summary>
            /// Apply a builtin cell style to this Column and its cells. The necessary CellStyle will be created if it doesn't yet exist in the StyleSheet.
            /// </summary>
            /// <param name="column"></param>
            /// <param name="style"></param>
            public static void ApplyCellStyle(this Column column, BuiltinCellStyle style)
            {
                for (uint i = column.Min; i <= column.Max; i++)
                {
                    new Range(column.Ancestors<Worksheet>().Single(), null, i).ApplyCellStyle(style);
                }
            }
        }

        public static class CellExtension
        {
            /// <summary>
            /// Apply the StyleIndex corresponding to the given CellFormat to this Cell. The CellFormat will be created if it doesn't yet exist in the StyleSheet.
            /// </summary>
            /// <param name="cell"></param>
            /// <param name="cellFormat"></param>
            public static void ApplyCellFormat(this Cell cell, CellFormat cellFormat)
            {
                cell.StyleIndex = cellFormat.GetNodeIndex((cell.Ancestors<Worksheet>().Single().WorksheetPart.GetParentParts().Single() as WorkbookPart).WorkbookStylesPart.Stylesheet.Elements<CellFormats>().Single());
            }

            /// <summary>
            /// Apply a builtin cell style to this Cell. The necessary CellStyle will be created if it doesn't yet exist in the StyleSheet.
            /// </summary>
            /// <param name="cell"></param>
            /// <param name="style"></param>
            public static void ApplyCellStyle(this Cell cell, BuiltinCellStyle style)
            {
                SpreadsheetDocument document = (SpreadsheetDocument)cell.Ancestors<Worksheet>().Single().WorksheetPart.OpenXmlPackage;
                CellFormat cellFormat = document.CreateStyleElementTemplate<CellFormat>(cell.StyleIndex);

                document.ApplyCellStyle(cellFormat, style);
                cell.ApplyCellFormat(cellFormat);
            }
        }
    }
    

    
}
