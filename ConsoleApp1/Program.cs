using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace ConsoleApp1
{
    class Program
    {
        static DataTable dt = new DataTable();

        static void Main(string[] args)
        {
            var filePath = @"E:\Bugzilla issue20200217.xlsx";
            var sheetName = "出貨日期";

            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value = sheets.First(x => x.Name == sheetName).Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                Row[] rows = sheetData.Descendants<Row>().ToArray();
                // 設置表頭DataTable
                foreach (Cell cell in rows.ElementAt(0))
                {
                    dt.Columns.Add((string)GetCellValue(spreadSheetDocument, cell));
                }
                // 內容
                for (int rowIndex = 1; rowIndex < rows.Count(); rowIndex++)
                {
                    DataRow tempRow = dt.NewRow();
                    for (int i = 0; i < rows[rowIndex].Descendants<Cell>().Count(); i++)
                    {
                        tempRow[i] = GetCellValue(spreadSheetDocument, rows[rowIndex].Descendants<Cell>().ElementAt(i));
                    }
                    dt.Rows.Add(tempRow);
                }
            }
            Save();
        }

        private static void Save()
        {
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();

                builder.DataSource = "127.0.0.1";
                builder.UserID = "sa";
                builder.Password = "p@ssw0rd";
                builder.InitialCatalog = "AF";

                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();

                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                    {
                        bulkCopy.ColumnMappings.Add("ROBOT SN ", "RobotSn");
                        bulkCopy.ColumnMappings.Add("HW Series", "HWSeries");


                        bulkCopy.DestinationTableName = "出貨日期";
                        try
                        {
                            bulkCopy.WriteToServer(dt);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                    connection.Close();
                }
            }
            catch (SqlException e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            
            if (cell.DataType != null && (cell.DataType.Value == CellValues.SharedString || cell.DataType.Value == CellValues.String || cell.DataType.Value == CellValues.Number))
            {
                string value = cell.CellValue.InnerXml;

                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else //浮點數和日期對應的cell.DataType都為NULL
            {
                // DateTime.FromOADate((double.Parse(value)); 如果確定是日期就可以直接用過該方法轉換為日期對象，可是無法確定DataType==NULL的時候這個CELL 數據到底是浮點型還是日期.(日期被自動轉換為浮點
                //return value;
                return "";

            }
        }
    }
}
