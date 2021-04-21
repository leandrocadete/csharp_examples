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
            ReadExcel(args[0]);
        }
        public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id) {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }
        public static void ReadExcel(string path) {

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, true)) {
                Sheets sheets = doc.WorkbookPart.Workbook.Sheets;
                WorkbookPart wPart = doc.WorkbookPart;

                IEnumerable<Sheet> shs = sheets.ChildElements.Cast<Sheet>();
                var sheet1 = shs.FirstOrDefault<Sheet>(s => s.Name == "2021");

                Worksheet workSheet = ((WorksheetPart)wPart.GetPartById(sheet1.Id)).Worksheet;
                List<SheetData> rows = workSheet.ChildElements.OfType<SheetData>().ToList();

                string currentcellvalue = string.Empty;

                for (int i = 0; i < rows[0].ChildElements.Count; i++) {
                    Row currentrow = (Row)rows[0].ChildElements.GetItem(i);

                    Cell[] cells = new Cell[] {
                        (Cell)currentrow.ChildElements.GetItem(0),
                        (Cell)currentrow.ChildElements.GetItem(1),
                        (Cell)currentrow.ChildElements.GetItem(2)
                    };

                    foreach (Cell c in cells) {
                        if(c.DataType != null) {
                            if(c.DataType == CellValues.SharedString) {
                                int id = -1;
                                if(int.TryParse(c.InnerText, out id)) {
                                    SharedStringItem item = GetSharedStringItemById(wPart, id);
                                    if(item.Text != null) {
                                        currentcellvalue = item.Text.Text;
                                    } else if (item.InnerText != null) {
                                        currentcellvalue = item.InnerText;
                                    } else if (item.InnerXml != null) {
                                        currentcellvalue = item.InnerXml;
                                    }
                                }
                            }
                        }
                        Console.Write("{0}; ", currentcellvalue);
                    }
                        Console.WriteLine();
                    


                }





            }
        }
    }
}
