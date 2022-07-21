using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SamuraiTools.OpenXml.Spreadsheet
{
    /// <summary>
    /// Collection representing a consecutive list of cells, a row, or a column.
    /// </summary>
    public class Range
    {
        protected enum RangeType
        {
            Cell = 0,
            Row = 1,
            Column = 2
        }

        protected RangeType Type { get; }

        protected List<Cell> Cells { get; } = new List<Cell>();
        protected Row Row { get; set; } = null;
        protected Column Column { get; set; } = null;

        public SpreadsheetDocument SpreadsheetDocument
        {
            get
            {
                return Worksheet.WorksheetPart?.OpenXmlPackage as SpreadsheetDocument;
            }
        }
        
        public Worksheet Worksheet { get; }
        public uint? StartRow { get; }
        public uint? StartColumn { get; }
        public uint? EndRow { get; }
        public uint? EndColumn { get; }
        public uint? RowCount { get { return EndRow - StartRow + 1U; } }
        public uint? ColumnCount { get { return EndColumn - StartColumn + 1U; } }

        protected int TotalCells
        {
            get
            {
                switch (Type)
                {
                    case RangeType.Cell:
                        return (int)(RowCount * ColumnCount);
                    case RangeType.Row:
                        return Row?.ChildElements.Count ?? 0;
                    case RangeType.Column:
                        int total = 0;
                        string columnLetter = SpreadsheetUtility.GetColumnLetter(StartColumn.Value);
                        foreach (var row in Worksheet.GetFirstChild<SheetData>().Elements<Row>())
                        {
                            total += row.Elements<Cell>().Where(c => SpreadsheetUtility.GetColumnLetter(c.CellReference) == columnLetter).Count();
                        }
                        return total;
                    default:
                        return -1;
                }
            }
        }
        protected bool IsMissingCells { get { return Cells.Count < TotalCells; } }

        protected Range()
        { }

        public Range(Worksheet worksheet, uint? startRow, uint? startColumn, uint? rowCount, uint? columnCount) : this()
        {
            if ((startRow.HasValue ^ rowCount.HasValue) || (startColumn.HasValue ^ columnCount.HasValue) || (!startRow.HasValue && !startColumn.HasValue) || 
                (startRow.HasValue && startRow == 0) || (startColumn.HasValue && startColumn == 0) || (rowCount.HasValue && rowCount == 0) || (columnCount.HasValue && columnCount == 0))
            {
                throw new ArgumentException("Invalid arguments when creating new Range");
            }

            StartRow = startRow;
            StartColumn = startColumn;
            EndRow = startRow + rowCount - 1;
            EndColumn = startColumn + columnCount - 1;
            Worksheet = worksheet;

            if (RowCount == null)
            {
                Type = RangeType.Column;
            }
            else if (ColumnCount == null)
            {
                Type = RangeType.Row;
            }
            else
            {
                Type = RangeType.Cell;
            }
        }

        public Range(Worksheet worksheet, uint? rowIndex, uint? columnIndex) : this(worksheet, rowIndex, columnIndex, rowIndex == null ? (uint?)null : 1U, columnIndex == null ? (uint?)null : 1U) { }

        /// <summary>
        /// Fill internal list of cells, row, and column in preparation for manipulating them in some way.
        /// Elements should not be written to the document unecessarily, so wait to call this until actually doing something with the Range.
        /// </summary>
        protected void PrepRangeElements()
        {
            switch (Type)
            {
                case RangeType.Cell:
                    if (IsMissingCells)
                    {
                        Cells.Clear();
                        for (uint i = StartRow.Value; i <= EndRow.Value; i++)
                        {
                            Row row = Worksheet.Rows(i);
                            for (uint j = StartColumn.Value; j <= EndColumn.Value; j++)
                            {
                                Cells.Add(row.Cells(j));
                            }
                        }
                    }
                    break;

                case RangeType.Column:
                    Column ??= Worksheet.Columns(StartColumn.Value);

                    if (IsMissingCells)
                    {
                        Cells.Clear();
                        string columnLetter = SpreadsheetUtility.GetColumnLetter(StartColumn.Value);
                        foreach (var row in Worksheet.GetFirstChild<SheetData>().Elements<Row>())
                        {
                            Cells.AddRange(row.Elements<Cell>().Where(c => SpreadsheetUtility.GetColumnLetter(c.CellReference) == columnLetter));
                        }
                    }
                    break;

                case RangeType.Row:
                    Row ??= Worksheet.Rows(StartRow.Value);

                    if (IsMissingCells)
                    {
                        Cells.Clear();
                        Cells.AddRange(Row.Elements<Cell>());
                    }
                    break;
            }
        }

        /// <summary>
        /// Merge the cells in this Range, optionally centering them. Only applicable to Cell ranges.
        /// </summary>
        /// <param name="center">Set the horizontal alignment to center after merging.</param>
        public void MergeCells(bool center)
        {
            if (Type != RangeType.Cell || TotalCells < 2)
            {
                return;
            }

            PrepRangeElements();

            string cellReference = Cells[0].CellReference.Value + ":" + Cells[Cells.Count - 1].CellReference.Value;
            MergeCells mergeCells = Worksheet.Elements<MergeCells>().FirstOrDefault() ?? Worksheet.InsertAfter(new MergeCells() { Count = 0 }, Worksheet.Elements<SheetData>().Single());

            if (mergeCells.Elements<MergeCell>().FirstOrDefault(c => c.Reference.Value == cellReference) == null)
            {
                mergeCells.AppendChild(new MergeCell() { Reference = cellReference });
                mergeCells.Count++;
            }

            if (center)
            {
                CellFormat referenceCellFormat = SpreadsheetDocument.CreateStyleElementTemplate<CellFormat>(Cells[0].StyleIndex);

                if (referenceCellFormat.Alignment == null)
                {
                    referenceCellFormat.Alignment = new Alignment();
                }

                referenceCellFormat.Alignment.Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Center);
                referenceCellFormat.ApplyAlignment = true;

                Cells[0].ApplyCellFormat(referenceCellFormat);

                for (int i = 1; i < Cells.Count; i++)
                {
                    Cells[i].StyleIndex = Cells[0].StyleIndex;
                }
            }

            for (int i = 1; i < Cells.Count; i++)
            {
                Cells[i].RemoveAllChildren();
            }
        }

        /// <summary>
        /// Auto fit width of column in the range. Not applicable to other range types.
        /// </summary>
        public void AutoFit()
        {
            if (Type == RangeType.Column)
            {
                (Column ?? Worksheet.Columns(StartColumn.Value)).AutoFit();
            }
        }

        /// <summary>
        /// Apply a builtin cell style to the cells in this Range. The necessary CellStyle will be created if it doesn't yet exist in the StyleSheet.
        /// </summary>
        /// <param name="style"></param>
        public void ApplyCellStyle(BuiltinCellStyle style)
        {
            PrepRangeElements();

            foreach (var cell in Cells)
            {
                cell.ApplyCellStyle(style);
            }
            if (Row != null)
            {
                CellFormat cellFormat = SpreadsheetDocument.CreateStyleElementTemplate<CellFormat>(Row.StyleIndex);

                SpreadsheetDocument.ApplyCellStyle(cellFormat, style);
                Row.StyleIndex = cellFormat.GetNodeIndex(SpreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>().Single());
            }
            if (Column != null)
            {
                CellFormat cellFormat = SpreadsheetDocument.CreateStyleElementTemplate<CellFormat>(Column.Style);

                SpreadsheetDocument.ApplyCellStyle(cellFormat, style);
                Column.Style = cellFormat.GetNodeIndex(SpreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>().Single());
            }
        }

        /// <summary>
        /// Apply a style index to each element in the range.
        /// </summary>
        /// <param name="ModifyElement">The function to modify a template of the style element being changed (e.g. Font) if the range element's current style index is not shared with any previous element in the range.</param>
        protected void ApplyStyle<T>(System.Action<T> ModifyElement) where T : OpenXmlElement
        {
            Dictionary<uint, uint> styleIndexPairs = new Dictionary<uint, uint>();
            //dictionaries can't have a null key, so keep track of what would be the null key separately.
            uint? nullIndex = null;

            IEnumerable<OpenXmlElement> elements = Cells;
            if (Row != null)
            {
                elements = elements.Append(Row);
            }
            if (Column != null)
            {
                elements = elements.Append(Column);
            }

            Func<OpenXmlElement, UInt32Value> GetStyleIndex = delegate (OpenXmlElement el)
            {
                if (el is Cell)
                {
                    return ((Cell)el).StyleIndex;
                }
                else if (el is Row)
                {
                    return ((Row)el).StyleIndex;
                }
                else
                {
                    return ((Column)el).Style;
                }
            };

            Action<OpenXmlElement, UInt32Value> SetStyleIndex = delegate (OpenXmlElement el, UInt32Value index)
            {
                if (el is Cell)
                {
                    ((Cell)el).StyleIndex = index;
                }
                else if (el is Row)
                {
                    ((Row)el).StyleIndex = index;
                }
                else
                {
                    ((Column)el).Style = index;
                }
            };
            
            foreach (OpenXmlElement element in elements)
            {
                uint? currentStyleIndex = GetStyleIndex(element)?.Value;
                uint styleIndex = 0U;
                if ((currentStyleIndex == null && nullIndex != null) || (currentStyleIndex != null && styleIndexPairs.TryGetValue(currentStyleIndex.Value, out styleIndex)))
                {
                    //this element's style index is shared with a previous element in the range, so assign its new style index directly without creating a new style.
                    SetStyleIndex(element, currentStyleIndex == null ? nullIndex.Value : styleIndex);
                }
                else
                {
                    //need to create a new CellFormat and possibly a new style element.
                    CellFormat format = SpreadsheetDocument.CreateStyleElementTemplate<CellFormat>(currentStyleIndex);

                    //if T is a CellFormat we don't need to create a new style element; the CellFormat is being edited directly.
                    if (typeof(T) == typeof(CellFormat))
                    {
                        ModifyElement(format as T);
                    }
                    else
                    {
                        CreateStyleElement(currentStyleIndex, format, (dynamic)ModifyElement);
                    }

                    uint newStyleIndex = format.GetNodeIndex(SpreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>().Single());

                    SetStyleIndex(element, newStyleIndex);

                    if (currentStyleIndex == null)
                    {
                        nullIndex = newStyleIndex;
                    }
                    else
                    {
                        styleIndexPairs.Add(currentStyleIndex.Value, newStyleIndex);
                    }
                }
            }
        }

        /// <summary>
        /// For use by ApplyStyle to create a template of a style element.
        /// </summary>
        /// <param name="styleIndex"></param>
        /// <param name="formatTemplate"></param>
        /// <param name="ModifyElement"></param>
        protected void CreateStyleElement(uint? styleIndex, CellFormat formatTemplate, Action<Font> ModifyElement)
        {
            Font font = SpreadsheetDocument.CreateStyleElementTemplate<Font>(styleIndex);

            ModifyElement(font);

            formatTemplate.FontId = font.GetNodeIndex(SpreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<Fonts>().Single());
            formatTemplate.ApplyFont = true;
        }

        protected void CreateStyleElement(uint? styleIndex, CellFormat formatTemplate, Action<Fill> ModifyElement)
        {
            Fill fill = SpreadsheetDocument.CreateStyleElementTemplate<Fill>(styleIndex);

            ModifyElement(fill);

            formatTemplate.FillId = fill.GetNodeIndex(SpreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<Fills>().Single());
            formatTemplate.ApplyFill = true;
        }

        protected void CreateStyleElement(uint? styleIndex, CellFormat formatTemplate, Action<Border> ModifyElement)
        {
            Border border = SpreadsheetDocument.CreateStyleElementTemplate<Border>(styleIndex);

            ModifyElement(border);

            formatTemplate.BorderId = border.GetNodeIndex(SpreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<Borders>().Single());
            formatTemplate.ApplyBorder = true;
        }

        #region Font styling
        /// <summary>
        /// Set many common characteristics of font.
        /// </summary>
        /// <param name="name">Font family name.</param>
        /// <param name="size">Font size in points</param>
        /// <param name="bold"></param>
        /// <param name="underline"></param>
        /// <param name="italic"></param>
        /// <param name="strikethrough"></param>
        /// <param name="color">Hex binary 32-bit color value.</param>
        /// <remarks>Any part can be null to leave it unchanged.</remarks>
        public void SetFont(string name = null, double? size = null, bool? bold = null, bool? underline = null, bool? italic = null, bool? strikethrough = null, string color = null)
        {
            PrepRangeElements();

            ApplyStyle(delegate (Font font)
            {
                //CellStyle normalStyle = SpreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<CellStyles>().FirstOrDefault()?.Elements<CellStyle>().SingleOrDefault(s => s.BuiltinId == (uint)BuiltinCellStyle.Normal);
                //Font normalFont = normalStyle == null ? null : SpreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<Fonts>().Single().Elements<Font>().ElementAt((int)SpreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.Elements<CellStyleFormats>().Single().Elements<CellFormat>().ElementAt((int)normalStyle.FormatId.Value).FontId.Value);

                if (name != null)
                {
                    font.FontName = new FontName() { Val = name };
                }
                if (size != null)
                {
                    font.FontSize = new FontSize() { Val = size.Value };
                }
                if (bold != null)
                {
                    font.Bold = bold.Value ? new Bold() : null;
                }
                if (underline != null)
                {
                    font.Underline = underline.Value ? new Underline() : null;
                }
                if (italic != null)
                {
                    font.Italic = italic.Value ? new Italic() : null;
                }
                if (strikethrough != null)
                {
                    font.Strike = strikethrough.Value ? new Strike() : null;
                }
                if (color != null)
                {
                    font.Color = new Color() { Rgb = new HexBinaryValue(color) };
                }
            });
        }
        #endregion

        #region Fill styling
        /// <summary>
        /// Set Gradient fill for the elements in this Range.
        /// </summary>
        /// <param name="gradientFillType"></param>
        /// <param name="color1">Beginning color of the gradient transition.</param>
        /// <param name="color2">Ending color of the gradient transition.</param>
        /// <remarks>Any existing fill for these elements will be replaced.</remarks>
        public void SetGradientFill(GradientFillType gradientFillType, string color1, string color2)
        {
            PrepRangeElements();

            ApplyStyle(delegate (Fill fill)
            {
                fill.RemoveAllChildren();
                fill.ClearAllAttributes();
                fill.GradientFill = new GradientFill() { Type = new EnumValue<GradientValues>(GradientValues.Linear) };

                switch (gradientFillType)
                {
                    case GradientFillType.Vertical:
                        fill.GradientFill.Degree = 90D;
                        fill.GradientFill.Append(new GradientStop(new Color() { Rgb = new HexBinaryValue(color1) }) { Position = 0D });
                        fill.GradientFill.Append(new GradientStop(new Color() { Rgb = new HexBinaryValue(color2) }) { Position = 1D });
                        break;

                    case GradientFillType.VerticalCenter:
                        fill.GradientFill.Degree = 90D;
                        fill.GradientFill.Append(new GradientStop(new Color() { Rgb = new HexBinaryValue(color1) }) { Position = 0D });
                        fill.GradientFill.Append(new GradientStop(new Color() { Rgb = new HexBinaryValue(color2) }) { Position = 0.5D });
                        fill.GradientFill.Append(new GradientStop(new Color() { Rgb = new HexBinaryValue(color1) }) { Position = 1D });
                        break;

                    case GradientFillType.Horizontal:
                        fill.GradientFill.Degree = 0D;
                        fill.GradientFill.Append(new GradientStop(new Color() { Rgb = new HexBinaryValue(color1) }) { Position = 0D });
                        fill.GradientFill.Append(new GradientStop(new Color() { Rgb = new HexBinaryValue(color2) }) { Position = 1D });
                        break;

                    case GradientFillType.HorizontalCenter:
                        fill.GradientFill.Degree = 0D;
                        fill.GradientFill.Append(new GradientStop(new Color() { Rgb = new HexBinaryValue(color1) }) { Position = 0D });
                        fill.GradientFill.Append(new GradientStop(new Color() { Rgb = new HexBinaryValue(color2) }) { Position = 0.5D });
                        fill.GradientFill.Append(new GradientStop(new Color() { Rgb = new HexBinaryValue(color1) }) { Position = 1D });
                        break;
                }
            });
        }

        /// <summary>
        /// Set Pattern fill for the elements in this Range.
        /// </summary>
        /// <param name="patternType"></param>
        /// <param name="foregroundColor"></param>
        /// <param name="backgroundColor"></param>
        /// <remarks>Any existing fill for these elements will be replaced. Foreground and background color are not both required, so either can be null or empty to avoid setting one.</remarks>
        public void SetPatternFill(PatternValues patternType, string foregroundColor, string backgroundColor)
        {
            PrepRangeElements();

            ApplyStyle(delegate (Fill fill)
            {
                fill.RemoveAllChildren();
                fill.ClearAllAttributes();
                fill.PatternFill = new PatternFill() { PatternType = new EnumValue<PatternValues>(patternType) };

                if (foregroundColor != null)
                {
                    fill.PatternFill.ForegroundColor = new ForegroundColor() { Rgb = new HexBinaryValue(foregroundColor) };
                }
                if (backgroundColor != null)
                {
                    fill.PatternFill.BackgroundColor = new BackgroundColor() { Rgb = new HexBinaryValue(backgroundColor) };
                }
            });
        }
        #endregion

        #region Border styling
        /// <summary>
        /// Set border styling and color for the elements in this Range.
        /// </summary>
        /// <param name="leftBorderStyle"></param>
        /// <param name="leftBorderColor"></param>
        /// <param name="rightBorderStyle"></param>
        /// <param name="rightBorderColor"></param>
        /// <param name="topBorderStyle"></param>
        /// <param name="topBorderColor"></param>
        /// <param name="bottomBorderStyle"></param>
        /// <param name="bottomBorderColor"></param>
        /// <param name="diagonalBorderStyle"></param>
        /// <param name="diagonalBorderColor"></param>
        /// <remarks>Colors are in hex binary 32-bit format. Any part can remain null to leave it unchanged.</remarks>
        public void SetBorder(BorderStyleValues? leftBorderStyle = null, string leftBorderColor = null, BorderStyleValues? rightBorderStyle = null, string rightBorderColor = null, BorderStyleValues? topBorderStyle = null, string topBorderColor = null, BorderStyleValues? bottomBorderStyle = null, string bottomBorderColor = null, BorderStyleValues? diagonalBorderStyle = null, string diagonalBorderColor = null, bool? diagonalUp = null, bool? diagonalDown = null)
        {
            PrepRangeElements();

            ApplyStyle(delegate (Border border)
            {
                if (leftBorderStyle != null)
                {
                    border.LeftBorder ??= new LeftBorder();
                    border.LeftBorder.Style = new EnumValue<BorderStyleValues>(leftBorderStyle.Value);
                }
                if (leftBorderColor != null)
                {
                    border.LeftBorder ??= new LeftBorder();
                    border.LeftBorder.Color = new Color() { Rgb = new HexBinaryValue(leftBorderColor) };
                }
                if (rightBorderStyle != null)
                {
                    border.RightBorder ??= new RightBorder();
                    border.RightBorder.Style = new EnumValue<BorderStyleValues>(rightBorderStyle.Value);
                }
                if (rightBorderColor != null)
                {
                    border.RightBorder ??= new RightBorder();
                    border.RightBorder.Color = new Color() { Rgb = new HexBinaryValue(rightBorderColor) };
                }
                if (topBorderStyle != null)
                {
                    border.TopBorder ??= new TopBorder();
                    border.TopBorder.Style = new EnumValue<BorderStyleValues>(topBorderStyle.Value);
                }
                if (topBorderColor != null)
                {
                    border.TopBorder ??= new TopBorder();
                    border.TopBorder.Color = new Color() { Rgb = new HexBinaryValue(topBorderColor) };
                }
                if (bottomBorderStyle != null)
                {
                    border.BottomBorder ??= new BottomBorder();
                    border.BottomBorder.Style = new EnumValue<BorderStyleValues>(bottomBorderStyle.Value);
                }
                if (bottomBorderColor != null)
                {
                    border.BottomBorder ??= new BottomBorder();
                    border.BottomBorder.Color = new Color() { Rgb = new HexBinaryValue(bottomBorderColor) };
                }
                if (diagonalBorderStyle != null)
                {
                    border.DiagonalBorder ??= new DiagonalBorder();
                    border.DiagonalBorder.Style = new EnumValue<BorderStyleValues>(diagonalBorderStyle.Value);
                }
                if (diagonalBorderColor != null)
                {
                    border.DiagonalBorder ??= new DiagonalBorder();
                    border.DiagonalBorder.Color = new Color() { Rgb = new HexBinaryValue(diagonalBorderColor) };
                }
                if (diagonalUp != null)
                {
                    border.DiagonalUp = diagonalUp.Value;
                }
                if (diagonalDown != null)
                {
                    border.DiagonalDown = diagonalDown.Value;
                }
            });
        }

        /// <summary>
        /// Set top, left, right, and bottom borders to the same style and/or color with optional diagonal border specified separately.
        /// </summary>
        /// <param name="outlineBorderStyle">Border style for top, left, right, and bottom borders.</param>
        /// <param name="outlineBorderColor">Color for top, left, right, and bottom borders.</param>
        /// <param name="diagonalBorderStyle"></param>
        /// <param name="diagonalBorderColor"></param>
        /// <param name="diagonalUp"></param>
        /// <param name="diagonalDown"></param>
        /// <remarks>Colors are in hex binary 32-bit format. Any part can remain null to leave it unchanged.</remarks>
        public void SetOutlineBorder(BorderStyleValues? outlineBorderStyle, string outlineBorderColor, BorderStyleValues? diagonalBorderStyle = null, string diagonalBorderColor = null, bool? diagonalUp = null, bool? diagonalDown = null)
        {
            SetBorder(outlineBorderStyle, outlineBorderColor, outlineBorderStyle, outlineBorderColor, outlineBorderStyle, outlineBorderColor, outlineBorderStyle, outlineBorderColor, diagonalBorderStyle, diagonalBorderColor, diagonalUp, diagonalDown);
        }

        #endregion

        /// <summary>
        /// Set alignment for elements in this range.
        /// </summary>
        /// <param name="horizontalAlignment"></param>
        /// <param name="verticalAlignment"></param>
        /// <param name="indent"></param>
        /// <param name="wrap"></param>
        /// <remarks> Any part can remain null to leave it unchanged.</remarks>
        public void SetAlignment(HorizontalAlignmentValues? horizontalAlignment = null, VerticalAlignmentValues? verticalAlignment = null, int? indent = null, bool? wrap = null)
        {
            PrepRangeElements();

            ApplyStyle(delegate (CellFormat formatTemplate)
            {
                if (horizontalAlignment != null)
                {
                    formatTemplate.Alignment ??= new Alignment();
                    formatTemplate.Alignment.Horizontal = new EnumValue<HorizontalAlignmentValues>(horizontalAlignment);
                }
                if (verticalAlignment != null)
                {
                    formatTemplate.Alignment ??= new Alignment();
                    formatTemplate.Alignment.Vertical = new EnumValue<VerticalAlignmentValues>(verticalAlignment);
                }
                if (indent != null)
                {
                    formatTemplate.Alignment ??= new Alignment();
                    formatTemplate.Alignment.Indent = (uint)indent;
                }
                if (wrap != null)
                {
                    formatTemplate.Alignment ??= new Alignment();
                    formatTemplate.Alignment.WrapText = new BooleanValue(wrap);
                }

                formatTemplate.ApplyAlignment = true;
            });
        }
        
        protected void SetValue(Row row, uint columnIndex, string value)
        {
            Cell cell = row.Cells(columnIndex);

            if (SpreadsheetDocument.WorkbookPart.SharedStringTablePart == null)
            {
                SpreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>().SharedStringTable = new SharedStringTable() { Count = 0U };
            }

            uint stringIndex = OpenXmlUtility.GetNodeIndex(new SharedStringItem(new Text(value)), SpreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable);
            uint oldStringIndex = 0;
            bool checkRemove = false;

            //if the cell already had a string, need to check if it should be removed from the shared string table
            if (cell.DataType?.Value == CellValues.SharedString)
            {
                oldStringIndex = uint.Parse(cell.CellValue.Text);
                checkRemove = true;
            }

            cell.CellValue = new CellValue(stringIndex.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            if (checkRemove && stringIndex != oldStringIndex)
            {
                SpreadsheetDocument.RemoveSharedStringItem(oldStringIndex);
            }
        }

        protected void SetValue(Row row, uint columnIndex, decimal value)
        {
            Cell cell = row.Cells(columnIndex);
            uint oldStringIndex = 0;
            bool checkRemove = false;

            //if the cell already had a string, need to check if it should be removed from the shared string table
            if (cell.DataType?.Value == CellValues.SharedString)
            {
                oldStringIndex = uint.Parse(cell.CellValue.Text);
                checkRemove = true;
            }

            cell.CellValue = new CellValue(value.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);

            if (checkRemove)
            {
                SpreadsheetDocument.RemoveSharedStringItem(oldStringIndex);
            }
        }

        protected void SetValue(Row row, uint columnIndex, double value)
        {
            SetValue(row, columnIndex, (decimal)value);
        }

        protected void SetValue(Row row, uint columnIndex, float value)
        {
            SetValue(row, columnIndex, (decimal)value);
        }

        protected void SetValue(Row row, uint columnIndex, long value)
        {
            SetValue(row, columnIndex, (decimal)value);
        }

        protected void SetValue(Row row, uint columnIndex, ulong value)
        {
            SetValue(row, columnIndex, (decimal)value);
        }

        protected void SetValue(Row row, uint columnIndex, int value)
        {
            SetValue(row, columnIndex, (decimal)value);
        }

        protected void SetValue(Row row, uint columnIndex, uint value)
        {
            SetValue(row, columnIndex, (decimal)value);
        }

        protected void SetValue(Row row, uint columnIndex, short value)
        {
            SetValue(row, columnIndex, (decimal)value);
        }

        protected void SetValue(Row row, uint columnIndex, ushort value)
        {
            SetValue(row, columnIndex, (decimal)value);
        }

        /// <summary>
        /// Set values for the cells in this Range. All cells will contain the provided value.
        /// </summary>
        /// <param name="value"></param>
        public void SetValues(object value)
        {
            object[,] values = new object[RowCount ?? 1, ColumnCount ?? 1];
            for (int i = 0; i < values.GetLength(0); i++)
            {
                for (int j = 0; j < values.GetLength(1); j++)
                {
                    values[i, j] = value;
                }
            }
            SetValues(values);
        }

        /// <summary>
        /// Set values for the cells in this Range. The provided values will be repeated for each row or column of cells.
        /// </summary>
        /// <param name="values"></param>
        public void SetValues(object[] values)
        {
            object[,] newValues = new object[RowCount ?? 1, values.GetLength(0)];
            for (int i = 0; i < newValues.GetLength(0); i++)
            {
                for (int j = 0; j < newValues.GetLength(1); j++)
                {
                    newValues[i, j] = values[j];
                }
            }
            SetValues(newValues);
        }
        
        /// <summary>
        /// Set values for the cells in this Range.
        /// </summary>
        /// <param name="values"></param>
        public void SetValues(object[,] values)
        {
            object[,] newValues;

            //values should be inverted for Column type ranges.
            if (Type == RangeType.Column)
            {
                newValues = new object[values.GetLength(1), Math.Min(values.GetLength(0), ColumnCount.Value)];
                for (int i = 0; i < newValues.GetLength(0); i++)
                {
                    for (int j = 0; j < newValues.GetLength(1); j++)
                    {
                        newValues[i, j] = values[j, i];
                    }
                }
            }
            else
            {
                newValues = new object[Math.Min(values.GetLength(0), RowCount.Value), Type == RangeType.Row ? values.GetLength(1) : Math.Min(values.GetLength(1), ColumnCount.Value)];
                for (int i = 0; i < newValues.GetLength(0); i++)
                {
                    for (int j = 0; j < newValues.GetLength(1); j++)
                    {
                        newValues[i, j] = values[i, j];
                    }
                }
            }
            
            for (int i = 0; i < newValues.GetLength(0); i++)
            {
                Row row = Worksheet.Rows((StartRow ?? 1) + (uint)i);
                for (int j = 0; j < newValues.GetLength(1); j++)
                {
                    if (newValues[i, j]?.ToString().Length > 0)
                    {
                        SetValue(row, (StartColumn ?? 1) + (uint)j, (dynamic)newValues[i, j]);
                    }
                }
            }
        }
    }
}
