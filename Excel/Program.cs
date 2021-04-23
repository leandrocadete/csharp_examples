using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace Excel {
    class Program {
        static void Main(string[] args) {

            ReadExcel(args[0]); // Get path from the first argument
        }
        
        public static void ReadExcel(string path) {

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, true)) {
                Sheets sheets = doc.WorkbookPart.Workbook.Sheets;
                WorkbookPart wPart = doc.WorkbookPart;

                IEnumerable<Sheet> shs = sheets.ChildElements.Cast<Sheet>();
                var sheet1 = shs.FirstOrDefault<Sheet>(s => s.Name == "2021");

                Worksheet workSheet = ((WorksheetPart)wPart.GetPartById(sheet1.Id)).Worksheet;
                List<SheetData> rows = workSheet.ChildElements.OfType<SheetData>().ToList();

                string currCellValue = string.Empty;

                List<List<string>> lstSheet = new List<List<string>>(rows[0].ChildElements.Count);

                for (int i = 0; i < rows[0].ChildElements.Count; i++) {
                    lstSheet.Add(new List<string>(4));


                    Row currentrow = (Row)rows[0].ChildElements.GetItem(i);

                    Cell[] cells = new Cell[] {
                        (Cell)currentrow.ChildElements.GetItem(0),
                        (Cell)currentrow.ChildElements.GetItem(1),
                        (Cell)currentrow.ChildElements.GetItem(2),
                        (Cell)currentrow.ChildElements.GetItem(3)
                    };

                    foreach (Cell c in cells) {
                        currCellValue = getStringFromCellValue(wPart, currCellValue, c);
                        lstSheet.Last().Add(currCellValue);
                    }
                }
                System.IO.StreamWriter strW = new System.IO.StreamWriter("test_output.csv");
                foreach (var rs in lstSheet) {
                    foreach (var c in rs) {
                        Console.Write("{0}; ", c);
                        strW.Write("{0};", c);
                    }
                    strW.WriteLine();
                    Console.WriteLine();
                }
                strW.Dispose();
                strW.Close();
            }
            
            
            // ------------------ Inner functions ------------------

            string getStringFromCellValue(WorkbookPart wPart, string currCellValue, Cell c) {
                if (c.DataType != null) {
                    Console.WriteLine("DataType: {0}", c.DataType.InnerText);
                    if (c.DataType == CellValues.SharedString) {
                        int id = -1;
                        if (int.TryParse(c.InnerText, out id)) {
                            SharedStringItem item = GetSharedStringItemById(wPart, id);
                            if (item.Text != null) {
                                currCellValue = item.Text.Text;
                            } else if (item.InnerText != null) {
                                currCellValue = item.InnerText;
                            } else if (item.InnerXml != null) {
                                currCellValue = item.InnerXml;
                            }
                        }
                    }
                } else {
                    Console.WriteLine("DataType: {0}", c.DataType?.InnerText);
                    currCellValue = c.InnerText;
                }

                return currCellValue;
            }
            SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id) {
                return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
            }

        }
    }
}
