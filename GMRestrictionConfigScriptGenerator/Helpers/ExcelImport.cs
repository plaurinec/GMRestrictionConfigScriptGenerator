using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace GMRestrictionConfigScriptGenerator.Helpers
{
    public class ExcelImport
    {
        /// <summary>
        /// Enumrate excel movements
        /// </summary>
        /// <param name="file"></param>
        /// <param name="create"></param>
        public void ImportMovementsFromExcel(string file, Action<MovementItem> create = null)
        {
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(file, false))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                // issues
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                SharedStringTable sst = null;
                var sstparts = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sstparts != null) sst = sstparts.SharedStringTable;
                
                foreach (var c in GetTableFromExcel(workbookPart, sheetData, sst, 1,0))
                {
                    if (create != null) create(c);
                };
                //recepits
                relationshipId = sheets.Skip(1).First().Id.Value;
                worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                workSheet = worksheetPart.Worksheet;
                sheetData = workSheet.GetFirstChild<SheetData>();
                sst = null;
                sstparts = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sstparts != null) sst = sstparts.SharedStringTable;
                foreach (var c in GetTableFromExcel(workbookPart, sheetData, sst, 0,0))
                {
                    if (create != null) create(c);
                };
            }
        }
        /// <summary>
        /// format datumu
        /// </summary>
        private readonly Dictionary<uint, string> DateFormatDictionary = new Dictionary<uint, string>()
        {
            [14] = "dd/MM/yyyy",
            [15] = "d-MMM-yy",
            [16] = "d-MMM",
            [17] = "MMM-yy",
            [18] = "h:mm AM/PM",
            [19] = "h:mm:ss AM/PM",
            [20] = "h:mm",
            [21] = "h:mm:ss",
            [22] = "M/d/yy h:mm",
            [30] = "M/d/yy",
            [34] = "yyyy-MM-dd",
            [45] = "mm:ss",
            [46] = "[h]:mm:ss",
            [47] = "mmss.0",
            [51] = "MM-dd",
            [52] = "yyyy-MM-dd",
            [53] = "yyyy-MM-dd",
            [55] = "yyyy-MM-dd",
            [56] = "yyyy-MM-dd",
            [58] = "MM-dd",
            [165] = "M/d/yy",
            [166] = "dd MMMM yyyy",
            [167] = "dd/MM/yyyy",
            [168] = "dd/MM/yy",
            [169] = "d.M.yy",
            [170] = "yyyy-MM-dd",
            [171] = "dd MMMM yyyy",
            [172] = "d MMMM yyyy",
            [173] = "M/d",
            [174] = "M/d/yy",
            [175] = "MM/dd/yy",
            [176] = "d-MMM",
            [177] = "d-MMM-yy",
            [178] = "dd-MMM-yy",
            [179] = "MMM-yy",
            [180] = "MMMM-yy",
            [181] = "MMMM d, yyyy",
            [182] = "M/d/yy hh:mm t",
            [183] = "M/d/y HH:mm",
            [184] = "MMM",
            [185] = "MMM-dd",
            [186] = "M/d/yyyy",
            [187] = "d-MMM-yyyy"
        };
        private string GetDateTimeFormat(UInt32Value numberFormatId)
        {
            return DateFormatDictionary.ContainsKey(numberFormatId) ? DateFormatDictionary[numberFormatId] : string.Empty;
        }
        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        /// <param name="cellReference">Address of the cell (ie. B2)</param>
        /// <returns>Column Name (ie. B)</returns>
        public static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        }
        /// <summary>
        /// Given just the column name (no row index), it will return the zero based column index.
        /// Note: This method will only handle columns with a length of up to two (ie. A to Z and AA to ZZ). 
        /// A length of three can be implemented when needed.
        /// </summary>
        /// <param name="columnName">Column Name (ie. A or AB)</param>
        /// <returns>Zero based index if the conversion was successful; otherwise null</returns>
        public static int? GetColumnIndexFromName(string columnName)
        {

            //return columnIndex;
            string name = columnName;
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }
            return number;
        }
        /// <summary>
        /// nacitanie hodnot z buniek
        /// </summary>
        /// <param name="row"></param>
        /// <param name="sst"></param>
        /// <returns></returns>
        private string GetValue(OpenXmlElement row, SharedStringTable sst)
        {
            var cell = row as Cell;
            if (cell == null) return string.Empty;
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                var value = sst.ChildElements[int.Parse(cell.CellValue.Text)].InnerText;
                return value;
            }
            return cell.InnerText;
        }

        private IEnumerable<MovementItem> GetTableFromExcel(WorkbookPart workbookPart, SheetData sheetData, SharedStringTable sst, int operation, int type)
        {
            var m = new List<ItemValue>();
            var rows = sheetData.Descendants<Row>().Skip(10);
            int columnIndex = 0;
            foreach (var r in rows)
            {
                var evi = new ItemValue();
                columnIndex = 0;
                foreach (Cell cell in r.Descendants<Cell>())
                {
                    string cellText = string.Empty;
                    int cellColumnIndex = (int)GetColumnIndexFromName(GetColumnName(cell.CellReference));
                    cellColumnIndex--; //zero based index
                    while (columnIndex < cellColumnIndex)
                    {
                        if (!SetItems(columnIndex, evi, cellText)) break;
                        columnIndex++;
                    }
                    columnIndex++;
                    var cellValue = cell.CellValue;
                    var text = (cellValue == null) ? cell.InnerText : cellValue.Text;
                    if (cell.DataType?.Value == CellValues.SharedString)
                    {
                        text = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Convert.ToInt32(cell.CellValue.Text)).InnerText;
                    }
                    cellText = (text ?? string.Empty).Trim();
                    if (cell.StyleIndex != null)
                    {
                        var cellFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[int.Parse(cell.StyleIndex.InnerText)] as CellFormat;
                        if (cellFormat != null)
                        {
                            var dateFormat = GetDateTimeFormat(cellFormat.NumberFormatId);
                            if (!string.IsNullOrEmpty(dateFormat))
                            {
                                if (double.TryParse(cellText, out var cellDouble))
                                {
                                    var theDate = DateTime.FromOADate(cellDouble);
                                    cellText = theDate.ToString("dd.MM.yyyy");
                                }
                            }
                        }
                    }
                    if (!SetItems(cellColumnIndex, evi, cellText) || cellColumnIndex == 1) break;
                }
                if (!string.IsNullOrWhiteSpace(evi.Prefix))
                {
                    m.Add(evi);
                }                
            }
            if (m.Count == 0) yield break;
            columnIndex = 2;
            while (true)
            {
                int rowIndex = 0;
                var ebi = new MovementItem() { Items = m.Select(p => new ItemValue() { Prefix = p.Prefix, Name = p.Name }).ToArray(), Operation = operation, Type = type };
                foreach (var r in sheetData.Descendants<Row>())
                {
                    foreach (Cell cell in r.Descendants<Cell>().Skip(columnIndex))
                    {
                        string cellText = string.Empty;
                        int cellColumnIndex = (int)GetColumnIndexFromName(GetColumnName(cell.CellReference));
                        cellColumnIndex--; //zero based index
                        if (columnIndex < cellColumnIndex)
                        {
                            SetMovementItem(rowIndex, ebi, cellText);
                            //columnIndex++;
                        }
                        //columnIndex++;
                        var cellValue = cell.CellValue;
                        var text = (cellValue == null) ? cell.InnerText : cellValue.Text;
                        if (cell.DataType?.Value == CellValues.SharedString)
                        {
                            text = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Convert.ToInt32(cell.CellValue.Text)).InnerText;
                        }
                        cellText = (text ?? string.Empty).Trim();
                        if (cell.StyleIndex != null)
                        {
                            var cellFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[int.Parse(cell.StyleIndex.InnerText)] as CellFormat;
                            if (cellFormat != null)
                            {
                                var dateFormat = GetDateTimeFormat(cellFormat.NumberFormatId);
                                if (!string.IsNullOrEmpty(dateFormat))
                                {
                                    if (double.TryParse(cellText, out var cellDouble))
                                    {
                                        var theDate = DateTime.FromOADate(cellDouble);
                                        cellText = theDate.ToString("dd.MM.yyyy");
                                    }
                                }
                            }
                        }
                        SetMovementItem(rowIndex, ebi, cellText);
                        break;
                    }
                    rowIndex++;
                    if (rowIndex > 1 && string.IsNullOrWhiteSpace(ebi.ExternalCode) && string.IsNullOrWhiteSpace(ebi.SubCode)) yield break;
                }
                //if (!string.IsNullOrWhiteSpace(evi.Prefix)) yield return ebi;
                yield return ebi;
                columnIndex++;
            }
        }
        private bool SetItems(int cellColumnIndex, ItemValue evi, string cellText)
        {
            if (cellColumnIndex > 0 && string.IsNullOrWhiteSpace(evi.Prefix)) return false;
            switch (cellColumnIndex)
            {
                case 0:
                    evi.Prefix = cellText.Trim();
                    break;
                case 1:
                    evi.Name = cellText.Trim();
                    break;   
            }
            return true;
        }
        private bool SetMovementItem(int rowIndex, MovementItem evi, string cellText)
        {
            switch (rowIndex)
            {
                case 0:
                    evi.ExternalCode = cellText.Trim();
                    break;
                case 1:
                    evi.SubCode = cellText.Trim();
                    break;
                case 2:
                    evi.ShortName = cellText.Trim();
                    break;
                case 3:
                case 4:
                case 5:
                case 6:
                case 7:
                case 8:
                    int i = rowIndex - 3;
                    bool k = new string[] { "1", "Y", "A", "P" }.ToList().IndexOf(cellText) != -1;
                    bool k2 = new string[] { "0", "N", "B" }.ToList().IndexOf(cellText) != -1;
                    if (k || k2)
                    {
                        if (k) evi.Limits[i] = true;
                        else evi.Limits[i] = false;
                    }
                    break;
                case 9:
                    evi.Name = cellText.Trim();
                    break;
                default:
                    i = rowIndex - 10;
                    k = new string[] { "1", "Y", "A", "P" }.ToList().IndexOf(cellText) != -1;
                    k2 = new string[] { "0", "N", "B" }.ToList().IndexOf(cellText) != -1;
                    if (!k && !k2) return false;
                    if (k) evi.Items[i].State = true;
                    else evi.Items[i].State = false;
                    break;            
            }
            return true;
        }
    }
    /// <summary>
    /// Item class
    /// </summary>
    public class ItemValue
    {
        public string Prefix { get; set; }
        public string Name { get; set; }
        public bool State { get; set; }
    }
    /// <summary>
    /// Movement class
    /// </summary>
    public class MovementItem
    {
        public int Operation { get; set; }
        public int Type { get; set; }
        public string ExternalCode { get; set; }
        public string SubCode { get; set; }
        public string Name { get; set; }
        public string ShortName { get; set; }
        public ItemValue[] Items { get; set; }

        public bool?[] Limits { get; set; } = new bool?[6];

        public bool IsExternalCode { get { return !string.IsNullOrEmpty(ExternalCode); } }
    }
}
