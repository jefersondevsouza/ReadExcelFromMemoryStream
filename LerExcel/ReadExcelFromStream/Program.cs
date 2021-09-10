using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelFromStream
{
    class Program
    {
        static void Main(string[] args)
        {
            //Is necessary install the DocumentFormat.OpenXml
            Read();
            Console.WriteLine("Press any key to finish! ");
            Console.ReadLine();
        }

        private static void Read()
        {
            string strDoc = @"C:\Users\temp\NOVO_.xlsx";

            using (Stream stream = File.Open(strDoc, FileMode.Open))
            {
                Dictionary<string, Dictionary<string, object>> dicSheet = MapSpreadsheetStream(stream);

                foreach (KeyValuePair<string, Dictionary<string, object>> sheetItemDic in dicSheet)
                {
                    foreach (KeyValuePair<string, object> cellItem in sheetItemDic.Value)
                    {
                        Console.WriteLine($"Sheet {sheetItemDic.Key} Cell = {cellItem.Key} = {cellItem.Value}");
                    }
                }
            }
        }

        public static Dictionary<string, Dictionary<string, object>> MapSpreadsheetStream(Stream stream)
        {
            var dicValores = new Dictionary<string, Dictionary<string, object>>();
            string value;

            using (var spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheets sheets = workbookPart.Workbook.Sheets;

                if (sheets?.Any() == true)
                {
                    foreach (Sheet sheet in sheets)
                    {
                        if (!dicValores.ContainsKey(sheet.Name))
                        {
                            dicValores.Add(sheet.Name, new Dictionary<string, object>());
                        }

                        var wsPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                        IEnumerable<Cell> cells = wsPart.Worksheet.Descendants<Cell>();

                        if (cells?.Any() == true)
                        {
                            foreach (Cell theCell in cells)
                            {
                                value = theCell.InnerText;

                                if (theCell.DataType != null)
                                {
                                    switch (theCell.DataType.Value)
                                    {
                                        case CellValues.SharedString:

                                            // For shared strings, look up the value in the
                                            // shared strings table.
                                            SharedStringTablePart stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                                            // If the shared string table is missing, something is wrong. Return the index that is in
                                            // the cell. Otherwise, look up the correct text in the table.
                                            if (stringTable != null)
                                            {
                                                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                                            }
                                            break;

                                        case CellValues.Boolean:
                                            switch (value)
                                            {
                                                case "0":
                                                    break;
                                                default:
                                                    value = "1";
                                                    break;
                                            }
                                            break;
                                    }
                                }

                                if (!dicValores[sheet.Name].ContainsKey(theCell.CellReference))
                                {
                                    dicValores[sheet.Name].Add(theCell.CellReference, value);
                                }
                            }
                        }
                    }
                }
            }

            return dicValores;
        }
    }
}
