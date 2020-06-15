using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace StoreMovementsScriptor.Helpers
{
    public class Spreadsheet
    {
        private static readonly Dictionary<string, int> m_shareStringDictionary = new Dictionary<string, int>();
        private static int m_shareStringMaxIndex = 0;

        /// <summary>
        /// Write xlsx spreadsheet file of a list of T objects
        /// Maximum of 24 columns
        /// </summary>
        /// <typeparam name="T">Type of objects passed in</typeparam>
        /// <param name="fileName">Full path filename for the new spreadsheet</param>
        /// <param name="def">A sheet definition used to create the spreadsheet</param>
        public static void Create<T>(string fileName, SheetDefinition<T>[] defs)
        {
            // open a template workbook
            using (var myWorkbook = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // create workbook part
                var workbookPart = myWorkbook.AddWorkbookPart();


                // add stylesheet to workbook part
                var stylesPart = myWorkbook.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                Stylesheet styles = new CustomStylesheet();
                //styles.Append(new CellFormats());
                //styles.CellFormats.Append(new CellFormat());
                //styles.CellFormats.Append(new CellFormat() { ApplyNumberFormat = true, NumberFormatId = 22 });
                styles.Save(stylesPart);

                //var cc = myWorkbook.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ToList();

                // create workbook
                var workbook = new Workbook();

                int sheetIndex = 1;
                var sheets = new Sheets();
                foreach (var def in defs)
                {
                    // add work sheet
                    sheets.AppendChild(CreateSheet(sheetIndex, def, workbookPart));                    
                    sheetIndex++;
                }
                workbook.AppendChild(sheets);

                CreateShareStringPart(myWorkbook);
                // add workbook to workbook part
                myWorkbook.WorkbookPart.Workbook = workbook;
                myWorkbook.WorkbookPart.Workbook.Save();
                myWorkbook.Close();
            }
        }

        private static Column CreateColumn(UInt32 startColumnIndex, UInt32 endColumnIndex, double width)
        {
            var column = new Column
            {
                Min = startColumnIndex,
                Max = endColumnIndex,
                BestFit = true,
                Width = width,
            };
            return column;
        }

        private static Sheet CreateSheet<T>(int sheetIndex, SheetDefinition<T> def, WorkbookPart workbookPart)
        {
            // create worksheet part
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var worksheetId = workbookPart.GetIdOfPart(worksheetPart);

            // variables
            var numCols = def.Fields.Count;
            //var numRows = def.Objects.Count;
            //var az = new List<Char>(Enumerable.Range('A', 'Z' - 'A' + 1).Select(i => (Char)i).ToArray());
            //var headerCols = az.GetRange(0, numCols);
            //var hasTitleRow = def.Title != null;
            //var hasSubtitleRow = def.SubTitle != null;
            //var titleRowCount = hasTitleRow ? 1 + (hasSubtitleRow ? 1 : 0) : hasSubtitleRow ? 1 : 0;

            // get the worksheet data
            //int firstTableRow;
            var sheetData = CreateSheetData(def.Objects, def.Fields, def.IncludeTotalsRow);//, def.Title, def.SubTitle, out firstTableRow);

            // populate column metadata
            var columns = new Columns();
            for (var col = 0; col < numCols; col++)
            {
                var width = ColumnWidth(sheetData, col, 0);// titleRowCount);
                columns.AppendChild(CreateColumn((UInt32)col + 1, (UInt32)col + 1, width));
            }

            // populate worksheet
            var worksheet = new Worksheet();
            //worksheet.AppendChild(columns);
            worksheet.AppendChild(sheetData);

            // add an auto filter
            //worksheet.AppendChild(new AutoFilter
            //{
            //    Reference =
            //        String.Format("{0}{1}:{2}{3}", headerCols.First(), firstTableRow - 1, headerCols.Last(),
            //            numRows + titleRowCount + 1)
            //});

            // add worksheet to worksheet part
            worksheetPart.Worksheet = worksheet;
            worksheetPart.Worksheet.Save();

            return new Sheet { Name = def.Name, SheetId = (UInt32)sheetIndex, Id = worksheetId };
        }

        private static double ColumnWidth(SheetData sheetData, int col, int fromRow)
        {
            //sheetData.ChildElements.getr
            double width = 8.43; //default width for most Excel installations
            try
            {
                var rows = sheetData.ChildElements.ToList().GetRange(fromRow, sheetData.ChildElements.Count - fromRow);
                int max = 0;
                foreach (var cc in rows) //.Where(p=>((Cell)p).CellReference.ToString().StartsWith(GetColumn(col))).Select(p => ((Cell)p.ChildElements[col]).InnerText))
                {
                    var cell = (Cell)cc.ChildElements.FirstOrDefault(p => ((Cell)p).CellReference.ToString().StartsWith(GetColumn(col)));
                    if (cell != null)
                    {
                        //var tf = new Typeface(new System.Windows.Media.FontFamily("Calibri"), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
                        //var ft = new FormattedText(cell.InnerText, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, tf, 11, Brushes.Black);
                        if (cell.InnerText.Length > max) max = cell.InnerText.Length;
                    }
                }
                //max = Math.Truncate(max / 4);
                width = Math.Truncate(((double)max * 7 + 5) / 7 * 256) / 256;
                width = Math.Round(width + 0.2, 2);
            }
            catch { }

            return width;
        }

        private static SheetData CreateSheetData<T>(IEnumerable<T> objects, List<SpreadsheetField> fields, bool includedTotalsRow) //string sheetTitle, string sheetSubTitle, , out int firstTableRow
        {
            var sheetData = new SheetData();
            var fieldNames = fields.Select(f => f.Title).ToList();
            //var numCols = headerCols.Count;
            var rowIndex = 0;
            //firstTableRow = 0;
            Row row;

            // create title
            //if (sheetTitle != null)
            //{
            //    rowIndex++;
            //    row = CreateTitle(sheetTitle, headerCols, ref rowIndex);
            //    sheetData.AppendChild(row);
            //}

            // create subtitle
            //if (sheetSubTitle != null)
            //{
            //    rowIndex++;
            //    row = CreateSubTitle(sheetSubTitle, headerCols, ref rowIndex);
            //    sheetData.AppendChild(row);
            //}

            // create the header
            //rowIndex++;
            //row = CreateHeader(fieldNames, ref rowIndex);
            //sheetData.AppendChild(row);

            //if (objects.Count == 0)
            //return sheetData;

            // create a row for each object and set the columns for each field
            //firstTableRow = rowIndex + 1;
            CreateTable(objects, ref rowIndex, fields, sheetData);

            // create an additional row with summed totals
            //if (includedTotalsRow)
            //{
            //    rowIndex++;
            //    AppendTotalsRow(objects, rowIndex, firstTableRow, numCols, fields, headerCols, sheetData);
            //}

            return sheetData;
        }

        private static Row CreateHeader(IList<string> headerNames, ref int rowIndex)
        {
            var header = new Row { RowIndex = (uint)rowIndex };

            for (var col = 0; col < headerNames.Count; col++)
            {
                var c = new TextCell(GetColumn(col), headerNames[col], rowIndex);
                //{
                //    StyleIndex = (UInt32)CustomStylesheet.CustomCellFormats.HeaderText;
                //};
                header.Append(c);
            }
            return header;
        }

        private static void CreateTable<T>(IEnumerable<T> objects, ref int rowIndex, List<SpreadsheetField> fields, SheetData sheetData, bool hidden = false, int outline = 0)
        {            
            // for each object
            foreach (var rowObj in objects)
            {
                if (rowObj is SheetRow)
                {
                    rowIndex++;

                    // create a row
                    var row = new Row
                    {
                        RowIndex = (uint)rowIndex,
                        Collapsed = new BooleanValue(false),
                        OutlineLevel = new ByteValue((byte)outline),
                        Hidden = new BooleanValue(hidden)
                    };
                    int col;
                    var raw = rowObj as SheetRow;
                    // populate columns using supplied objects
                    for (col = 0; col < raw.Columns.Count; col++)
                    {
                        //var field = fields[col];
                        //object columnObj = null;
                        //var columnObj = rowObj.GetObjectValue(field.FieldName);//  GetColumnObject(field.FieldName, rowObj);
                        var craw = raw.Columns[col];

                        if (string.IsNullOrEmpty(craw.Value)) continue;

                        Cell cell = null;

                        //if (field.CellType == CellValues.InlineString)
                        //{
                        //if (!m_shareStringDictionary.ContainsKey(craw.Value))
                        //{
                        //    m_shareStringDictionary.Add(craw.Value, m_shareStringMaxIndex);
                        //    m_shareStringMaxIndex++;
                        //}
                        //cell = new SharedTextCell(GetColumn(col), m_shareStringDictionary[craw.Value].ToString(), rowIndex);
                        cell = new TextCell(GetColumn(col), craw.Value, rowIndex);
                        if (craw.Bold)
                        {
                            cell.StyleIndex = (UInt32Value)1;// (UInt32)CustomStylesheet.CustomCellFormats.HeaderText;
                        };
                        //}                        
                        if (cell != null) row.AppendChild(cell);

                    } // for each column

                    sheetData.AppendChild(row);
                }
                else
                {
                    // row group?
                    var list = rowObj as IList<object>;
                    if (list != null)
                    {
                        CreateTable(list, ref rowIndex, fields, sheetData, true, outline + 1);
                        continue;
                    }

                    rowIndex++;

                    // create a row
                    var row = new Row
                    {
                        RowIndex = (uint)rowIndex,
                        Collapsed = new BooleanValue(false),
                        OutlineLevel = new ByteValue((byte)outline),
                        Hidden = new BooleanValue(hidden)
                    };

                    int col;

                    // populate columns using supplied objects
                    for (col = 0; col < fields.Count; col++)
                    {
                        var field = fields[col];
                        object columnObj = null;
                        //var columnObj = rowObj.GetObjectValue(field.FieldName);//  GetColumnObject(field.FieldName, rowObj);
                        if (columnObj == null || columnObj == DBNull.Value) continue;

                        Cell cell = null;

                        if (field.CellType == CellValues.InlineString)
                        {
                            if (!string.IsNullOrEmpty(columnObj.ToString()))
                            {
                                cell = new TextCell(GetColumn(col), columnObj.ToString(), rowIndex);
                            }
                        }
                        else if (field.CellType == CellValues.SharedString)
                        {
                            if (!string.IsNullOrEmpty(columnObj.ToString()))
                            {
                                if (!m_shareStringDictionary.ContainsKey(columnObj.ToString()))
                                {
                                    m_shareStringDictionary.Add(columnObj.ToString(), m_shareStringMaxIndex);
                                    m_shareStringMaxIndex++;
                                }
                                cell = new SharedTextCell(GetColumn(col), m_shareStringDictionary[columnObj.ToString()].ToString(), rowIndex);
                            }
                        }
                        else if (field.CellType == CellValues.Number)
                        {
                            cell = new NumberCell(GetColumn(col), (int)columnObj, rowIndex);
                        }
                        else if (field.CellType == CellValues.Date)
                        {
                            cell = new DateCell(GetColumn(col), (DateTime)columnObj, rowIndex);
                        }
                        else if (field.CellType == CellValues.Boolean)
                        {
                            cell = new BooleanCell(GetColumn(col), (bool)columnObj, rowIndex);
                        }
                        if (cell != null) row.AppendChild(cell);

                    } // for each column

                    sheetData.AppendChild(row);
                }
            }
        }

        private static void CreateShareStringPart(SpreadsheetDocument document)
        {
            if (m_shareStringMaxIndex > 0)
            {
                var sharedStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
                using (var writer = OpenXmlWriter.Create(sharedStringPart))
                {
                    writer.WriteStartElement(new SharedStringTable());
                    foreach (var item in m_shareStringDictionary)
                    {
                        writer.WriteStartElement(new SharedStringItem());
                        writer.WriteElement(new Text(item.Key));
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                }
            }
        }

        private static string GetColumn(int index)
        {
            return ((char)('A' + index)).ToString();
        }
    }
    //public class DateCellOld : Cell
    //{
    //    public DateCellOld(string header, DateTime dateTime, int index)
    //    {
    //        DataType = CellValues.Date;
    //        CellReference = header + index;
    //        double oaValue = dateTime.ToOADate();
    //        CellValue = new CellValue(oaValue.ToString(CultureInfo.InvariantCulture));
    //        DataType = new EnumValue<CellValues>(CellValues.Number);
    //    }
    //}
    public class DateCell : Cell
    {
        public DateCell(string header, DateTime dateTime, int index)
        {
            //DataType = CellValues.Date;
            CellReference = header + index;
            CellValue = new CellValue(dateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
            StyleIndex = 1;
            DataType = new EnumValue<CellValues>(CellValues.Number);
        }
    }
    public class NumberCell : Cell
    {
        public NumberCell(string header, int number, int index)
        {
            //DataType = CellValues.Date;
            CellReference = header + index;
            CellValue = new CellValue(number.ToString(CultureInfo.InvariantCulture));
            DataType = new EnumValue<CellValues>(CellValues.Number);
        }
    }
    public class TextCell : Cell
    {
        public TextCell(string header, string text, int index)
        {
            CellReference = header + index;
            InlineString = new InlineString { Text = new Text { Text = text } };
            DataType = new EnumValue<CellValues>(CellValues.InlineString);
        }
    }

    public class SharedTextCell : Cell
    {
        public SharedTextCell(string header, string text, int index)
        {
            //DataType = CellValues.Date;
            CellReference = header + index;
            CellValue = new CellValue(text);
            DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }
    }
    public class BooleanCell : Cell
    {
        public BooleanCell(string header, bool val, int index)
        {
            //DataType = CellValues.Date;
            CellReference = header + index;
            CellValue = new CellValue(val ? "1" : "0");
            DataType = new EnumValue<CellValues>(CellValues.Boolean);
        }
    }
    public class SpreadsheetField
    {
        public string Title { get; set; }
        public string FieldName { get; set; }
        public int DecimalPlaces { get; set; }
        public CellValues CellType { get; set; }
        //public bool IgnoreFromTotals { get; set; }
        //public bool CountNoneNullRowsForTotal { get; set; }
        public SpreadsheetField() { CellType = CellValues.SharedString; }
    }

    public class SheetDefinition<T>
    {
        /// <summary>
        /// Name of the sheet (shown in the tab)
        /// </summary>
        public string Name { get; set; }

        ///// <summary>
        ///// Title of the sheet
        ///// </summary>
        //public string Title { get; set; }

        ///// <summary>
        ///// Subtitle of the sheet
        ///// </summary>
        //public string SubTitle { get; set; }

        /// <summary>
        /// Objects to display in the sheet
        /// </summary>
        public IEnumerable<T> Objects { get; set; }

        /// <summary>
        /// Field names to extract from the objects and use as header names
        /// </summary>
        public List<SpreadsheetField> Fields { get; set; }

        /// <summary>
        /// Whether or not to include a row of calculated totals to the table
        /// </summary>
        public bool IncludeTotalsRow { get; set; }

    }
    public class CustomStylesheet : Stylesheet
    {
        public enum CustomCellFormats : uint
        {
            // these are referenced by index, must be added in this order
            DefaultText = 0,
            DefaultDate,
            DefaultNumber2DecimalPlace,
            DefaultNumber5DecimalPlace,
            DefaultDateTime,
            HeaderText,
            TotalsNumber,
            TotalsNumber2DecimalPlace,
            TotalsText,
            TitleText,
            SubtitleText,
            Duration,
            TotalsDuration,
            Hyperlink
        }

        public CustomStylesheet()
        {
            NumberingFormat nfDateTime;
            NumberingFormat nf5Decimal;
            NumberingFormat nfDuration;
            NumberingFormat nfTotalDuration;

            Append(CreateNumberingFormats(out nfDateTime, out nf5Decimal, out nfDuration, out nfTotalDuration));
            Append(CreateFonts());
            Append(CreateFills());
            Append(CreateBorders());
            Append(CreateCellStyleFormats());
            Append(CreateCellFormats(nfDateTime, nf5Decimal, nfDuration, nfTotalDuration));
            Append(CreateCellStyles());
            Append(CreateDifferentialFormats());
            //Append(CreateTableStyles());
        }

        //private static TableStyles CreateTableStyles()
        //{
        //    var tss = new TableStyles();
        //    tss.Count = 0;
        //    tss.DefaultTableStyle = StringValue.FromString("TableStyleMedium9");
        //    tss.DefaultPivotStyle = StringValue.FromString("PivotStyleLight16");
        //    return tss;
        //}

        private static DifferentialFormats CreateDifferentialFormats()
        {
            var dfs = new DifferentialFormats();
            dfs.Count = 0;
            return dfs;
        }

        private static CellStyles CreateCellStyles()
        {
            var css = new CellStyles();

            // cell style 0
            var cs = new CellStyle();
            cs.Name = StringValue.FromString("Normal");
            cs.FormatId = 0;
            cs.BuiltinId = 0;
            css.AppendChild(cs);
            css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
            return css;
        }

        /// <summary>
        /// Ensure cell formats are added in the order specified by the enumeration
        /// </summary>
        private static CellFormats CreateCellFormats(NumberingFormat nfDateTime, NumberingFormat nf5Decimal,
            NumberingFormat nfDuration, NumberingFormat nfTotalDuration)
        {
            var cfs = new CellFormats();

            // CustomCellFormats.DefaultText
            var cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(false);
            cfs.AppendChild(cf);

            cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 1;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(false);
            cfs.AppendChild(cf);

            // CustomCellFormats.DefaultDate
            cf = new CellFormat();
            cf.NumberFormatId = 22; // mm-dd-yy
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cfs.AppendChild(cf);

            // CustomCellFormats.DefaultNumber2DecimalPlace
            //cf = new CellFormat();
            //cf.NumberFormatId = 4; // #,##0.00
            //cf.FontId = 0;
            //cf.FillId = 0;
            //cf.BorderId = 0;
            //cf.FormatId = 0;
            //cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            //cfs.AppendChild(cf);

            //// CustomCellFormats.DefaultNumber5DecimalPlace
            //cf = new CellFormat();
            //cf.NumberFormatId = nf5Decimal.NumberFormatId;
            //cf.FontId = 0;
            //cf.FillId = 0;
            //cf.BorderId = 0;
            //cf.FormatId = 0;
            //cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            //cfs.AppendChild(cf);

            //// CustomCellFormats.DefaultDateTime
            //cf = new CellFormat();
            //cf.NumberFormatId = nfDateTime.NumberFormatId;
            //cf.FontId = 0;
            //cf.FillId = 0;
            //cf.BorderId = 0;
            //cf.FormatId = 0;
            //cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            //cfs.AppendChild(cf);

            //// CustomCellFormats.HeaderText
            //cf = new CellFormat();
            //cf.NumberFormatId = 0;
            //cf.FontId = 1;
            //cf.FillId = 2;
            //cf.BorderId = 0;
            //cf.FormatId = 0;
            //cf.ApplyNumberFormat = BooleanValue.FromBoolean(false);
            //cfs.AppendChild(cf);

            //// CustomCellFormats.TotalsNumber
            //cf = new CellFormat();
            //cf.NumberFormatId = 0;
            //cf.FontId = 0;
            //cf.FillId = 3;
            //cf.BorderId = 2;
            //cf.FormatId = 0;
            //cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            //cfs.AppendChild(cf);

            //// CustomCellFormats.TotalsNumber2DecimalPlace
            ////cf = new CellFormat();
            ////cf.NumberFormatId = 4; // #,##0.00
            ////cf.FontId = 0;
            ////cf.FillId = 3;
            ////cf.BorderId = 2;
            ////cf.FormatId = 0;
            ////cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            ////cfs.AppendChild(cf);

            //// CustomCellFormats.TotalsText
            ////cf = new CellFormat();
            ////cf.NumberFormatId = 49; // @
            ////cf.FontId = 0;
            ////cf.FillId = 3;
            ////cf.BorderId = 2;
            ////cf.FormatId = 0;
            ////cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            ////cfs.AppendChild(cf);

            //// CustomCellFormats.TitleText
            //cf = new CellFormat();
            //cf.NumberFormatId = 0;
            //cf.FontId = 2;
            //cf.FillId = 0;
            //cf.BorderId = 0;
            //cf.FormatId = 0;
            //cf.ApplyNumberFormat = BooleanValue.FromBoolean(false);
            //cf.Alignment = new Alignment
            //{
            //    Vertical = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Bottom)
            //};
            //cfs.AppendChild(cf);

            //// CustomCellFormats.SubtitleText
            //cf = new CellFormat();
            //cf.NumberFormatId = 0;
            //cf.FontId = 3;
            //cf.FillId = 0;
            //cf.BorderId = 0;
            //cf.FormatId = 0;
            //cf.ApplyNumberFormat = BooleanValue.FromBoolean(false);
            //cf.Alignment = new Alignment
            //{
            //    Vertical = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Top)
            //};
            //cfs.AppendChild(cf);

            //// CustomCellFormats.Duration
            //cf = new CellFormat();
            //cf.NumberFormatId = nfDuration.NumberFormatId; // [h]:mm
            //cf.FontId = 0;
            //cf.FillId = 0;
            //cf.BorderId = 0;
            //cf.FormatId = 0;
            //cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            //cf.Alignment = new Alignment
            //{
            //    Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Right)
            //};
            //cfs.AppendChild(cf);

            //// CustomCellFormats.TotalsNumber
            ////cf = new CellFormat();
            ////cf.NumberFormatId = nfTotalDuration.NumberFormatId; // d:h:mm
            ////cf.FontId = 0;
            ////cf.FillId = 3;
            ////cf.BorderId = 2;
            ////cf.FormatId = 0;
            ////cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            ////cf.Alignment = new Alignment
            ////{
            ////    Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Right)
            ////};
            ////cfs.AppendChild(cf);

            //// CustomCellFormats.Hyperlink
            //cf = new CellFormat();
            //cf.NumberFormatId = 0;
            //cf.FontId = 4;
            //cf.FillId = 0;
            //cf.BorderId = 0;
            //cf.FormatId = 0;
            //cf.ApplyNumberFormat = BooleanValue.FromBoolean(false);
            //cfs.AppendChild(cf);

            cfs.Count = UInt32Value.FromUInt32((uint)cfs.ChildElements.Count);
            return cfs;
        }

        private static NumberingFormats CreateNumberingFormats(out NumberingFormat nfDateTime,
            out NumberingFormat nf5Decimal, out NumberingFormat nfDuration, out NumberingFormat nfTotalDuration)
        {
            // built-in formats go up to 164
            uint iExcelIndex = 164;

            var nfs = new NumberingFormats();
            nfDateTime = new NumberingFormat();
            nfDateTime.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nfDateTime.FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss");
            nfs.AppendChild(nfDateTime);

            nf5Decimal = new NumberingFormat();
            nf5Decimal.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nf5Decimal.FormatCode = StringValue.FromString("#,##0.00000");
            nfs.AppendChild(nf5Decimal);

            nfDuration = new NumberingFormat();
            nfDuration.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nfDuration.FormatCode = StringValue.FromString("[h]:mm");
            nfs.AppendChild(nfDuration);

            nfTotalDuration = new NumberingFormat();
            nfTotalDuration.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nfTotalDuration.FormatCode = StringValue.FromString("d:h:mm");
            nfs.AppendChild(nfTotalDuration);

            nfs.Count = UInt32Value.FromUInt32((uint)nfs.ChildElements.Count);
            return nfs;
        }

        private static CellStyleFormats CreateCellStyleFormats()
        {
            var csfs = new CellStyleFormats();

            // cell style 0
            var cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            csfs.AppendChild(cf);
            csfs.Count = UInt32Value.FromUInt32((uint)csfs.ChildElements.Count);
            return csfs;
        }

        private static Borders CreateBorders()
        {
            var borders = new Borders();

            // boarder index 0
            var border = new Border();
            border.LeftBorder = new LeftBorder();
            border.RightBorder = new RightBorder();
            border.TopBorder = new TopBorder();
            border.BottomBorder = new BottomBorder();
            border.DiagonalBorder = new DiagonalBorder();
            borders.AppendChild(border);

            // boarder Index 1
            border = new Border();
            border.LeftBorder = new LeftBorder();
            border.LeftBorder.Style = BorderStyleValues.Thin;
            border.RightBorder = new RightBorder();
            border.RightBorder.Style = BorderStyleValues.Thin;
            border.TopBorder = new TopBorder();
            border.TopBorder.Style = BorderStyleValues.Thin;
            border.BottomBorder = new BottomBorder();
            border.BottomBorder.Style = BorderStyleValues.Thin;
            border.DiagonalBorder = new DiagonalBorder();
            borders.AppendChild(border);

            // boarder Index 2
            border = new Border();
            border.LeftBorder = new LeftBorder();
            border.RightBorder = new RightBorder();
            border.TopBorder = new TopBorder();
            border.TopBorder.Style = BorderStyleValues.Thin;
            border.BottomBorder = new BottomBorder();
            border.BottomBorder.Style = BorderStyleValues.Thin;
            border.DiagonalBorder = new DiagonalBorder();
            borders.AppendChild(border);

            borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);
            return borders;
        }

        private static Fills CreateFills()
        {
            // fill 0
            var fills = new Fills();
            var fill = new Fill();
            var patternFill = new PatternFill { PatternType = PatternValues.None };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            // fill 1 (in-built fill)
            fill = new Fill();
            patternFill = new PatternFill { PatternType = PatternValues.Gray125 };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            // fill 2
            //fill = new Fill();
            //patternFill = new PatternFill();
            //patternFill.PatternType = PatternValues.Solid;
            //var fillColor = Color.LightSkyBlue;
            //patternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValueFromColor(fillColor) };
            //patternFill.BackgroundColor = new BackgroundColor { Rgb = HexBinaryValueFromColor(fillColor) };
            //fill.PatternFill = patternFill;
            //fills.AppendChild(fill);

            //// fill 3
            //fill = new Fill();
            //patternFill = new PatternFill();
            //patternFill.PatternType = PatternValues.Solid;
            //fillColor = Color.Orange;
            //patternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValueFromColor(fillColor) };
            //patternFill.BackgroundColor = new BackgroundColor { Rgb = HexBinaryValueFromColor(fillColor) };
            //fill.PatternFill = patternFill;
            //fills.AppendChild(fill);

            fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);
            return fills;
        }

        private static Fonts CreateFonts()
        {
            var fts = new Fonts();

            // font 0
            var ft = new Font();
            var ftn = new FontName { Val = StringValue.FromString("Calibri") };
            var ftsz = new FontSize { Val = DoubleValue.FromDouble(11) };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            fts.AppendChild(ft);

            // font 1
            ft = new Font();
            ftn = new FontName { Val = StringValue.FromString("Calibri") };
            ftsz = new FontSize { Val = DoubleValue.FromDouble(11) };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            ft.Bold = new Bold();
            fts.AppendChild(ft);

            // font 2
            ft = new Font();
            ftn = new FontName { Val = StringValue.FromString("Calibri") };
            ftsz = new FontSize { Val = DoubleValue.FromDouble(18) };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            ft.Bold = new Bold();
            fts.AppendChild(ft);

            // font 3
            ft = new Font();
            ftn = new FontName { Val = StringValue.FromString("Calibri") };
            ftsz = new FontSize { Val = DoubleValue.FromDouble(14) };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            fts.AppendChild(ft);

            // font 4
            //ft = new Font();
            //ftn = new FontName { Val = StringValue.FromString("Arial") };
            //ftsz = new FontSize { Val = DoubleValue.FromDouble(11) };
            //var fontColor = Color.MediumBlue;
            //ft.Color = new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = HexBinaryValueFromColor(fontColor) };
            //ft.FontName = ftn;
            //ft.FontSize = ftsz;
            //fts.AppendChild(ft);

            fts.Count = UInt32Value.FromUInt32((uint)fts.ChildElements.Count);
            return fts;
        }

        //private static HexBinaryValue HexBinaryValueFromColor(Color fillColor)
        //{
        //    return new HexBinaryValue
        //    {
        //        Value =
        //            ColorTranslator.ToHtml(
        //                Color.FromArgb(
        //                    fillColor.A,
        //                    fillColor.R,
        //                    fillColor.G,
        //                    fillColor.B)).Replace("#", "")
        //    };
        //}
    }

    public class SheetRow
    {
        public List<SheetCellValue> Columns { get; set; } = new List<SheetCellValue>();
    }
    public class SheetCellValue
    {
        public string Value { get; set; }
        public bool Bold { get; set; }
    }

}
