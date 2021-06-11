using EPPlusJob.Options;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusJob.Actions
{
    public class DatasetToExcelWithFormattingAndRowStart
    {
        OptionOne options;
        Logger logger;
        public DatasetToExcelWithFormattingAndRowStart(OptionOne options, Logger _logger)
        {
            this.options = options;
            logger = _logger;
        }

        public void Run()
        {
            var dt = GetCollection();
            if(options.OutputType == outputType.Excel)
            {
                GenerateExcel(dt);
            }
            else if(options.OutputType == outputType.CSV)
            {
                var csvData = ConvertDataTableToCsvFile(dt);
                SaveData(csvData, options.FilePath);
            }
            
        }

        private DataTable GetCollection()
        {
            DataTable Results = new DataTable();
            logger.Info("Connecting to Database");
            using(SqlConnection connection = new SqlConnection(options.ConnectionString))
            {
                connection.Open();
                var dtA = new SqlDataAdapter(options.Query, connection);
                dtA.SelectCommand.CommandTimeout = 120;
                DataSet oDataSet = new DataSet();
                logger.Info("Executing Query");
                dtA.Fill(oDataSet);

                Results = oDataSet.Tables[0];
            }

            return Results;
        }
        private void GenerateExcel(DataTable Coll)
        {
           
            var file = Utility.GetFileInfo(options.FilePath, options.DeleteExistingFile);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage xlPackage = new ExcelPackage(file))
            {
                //ExcelWorksheet sheetcreate = xlPackage.Workbook.Worksheets.Add(SheetName);
                //ws.Cells[CellRef].LoadFromDataTable(Coll, true);
                logger.Info("Writing to Excel");
                ExcelWorksheet sheetcreate = xlPackage.Workbook.Worksheets[options.SheetName];
                if (sheetcreate == null)
                {
                    sheetcreate = xlPackage.Workbook.Worksheets.Add(options.SheetName);
                }

                int col = 0;
                int rowStart = Convert.ToInt32(options.RowStart);
                foreach (DataColumn column in Coll.Columns)  //printing column headings
                {
                    sheetcreate.Cells[rowStart, ++col].Value = column.ColumnName;
                    sheetcreate.Cells[rowStart, col].Style.Font.Bold = true;
                    sheetcreate.Cells[rowStart, col].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    sheetcreate.Cells[rowStart, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                if (Coll.Rows.Count > 0)
                {
                    int row = rowStart;
                    decimal checkDecimal;
                    DateTime checkDate;
                    for (int eachRow = 0; eachRow < Coll.Rows.Count;)    //looping each row
                    {
                        for (int eachColumn = 1; eachColumn <= col; eachColumn++)   //looping each column in a row
                        {
                            var eachRowObject = sheetcreate.Cells[row + 1, eachColumn];
                            eachRowObject.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            eachRowObject.Value = Coll.Rows[eachRow][(eachColumn - 1)].ToString();
                            if (decimal.TryParse(Coll.Rows[eachRow][(eachColumn - 1)].ToString(), out checkDecimal))      //verifying value is number
                            {
                                eachRowObject.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                eachRowObject.Style.Numberformat.Format = options.NumberFormat;
                            }
                            /*if (CheckDate) {
                                if (DateTime.TryParse(Coll.Rows[eachRow][(eachColumn - 1)].ToString(), out checkDate))      //verifying value is date 
                                {
                                    eachRowObject.Value = checkDate.Add(TimeCorrection); //Add 1 hour to datetime to fix BP issues
                                    eachRowObject.Style.Numberformat.Format = "yyyy-MM-ddTHH:mm:ss";
                                }
                            }*/
                            if (!string.IsNullOrEmpty(options.DateTimeColumns))
                            {
                                List<string> dtColumn = new List<string>();
                                dtColumn.AddRange(options.DateTimeColumns.Split(','));
                                if (eachColumn.ToString() == dtColumn.Where(x => Convert.ToInt32(x) == eachColumn).FirstOrDefault())
                                {
                                    if (DateTime.TryParse(Coll.Rows[eachRow][(eachColumn - 1)].ToString(), out checkDate))      //verifying value is date 
                                    {
                                        eachRowObject.Value = checkDate.Add(new TimeSpan(0,1,0,0)); //Add 1 hour to datetime to fix BP issues
                                    }
                                }
                            }
                            eachRowObject.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);     // adding border to each cells
                            if (eachRow % 2 == 0)       //alternatively adding color to each cell.
                                eachRowObject.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#e0e0e0"));
                            else
                                eachRowObject.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#ffffff"));
                        }
                        eachRow++;
                        row++;

                    }
                }
                sheetcreate.Cells.AutoFitColumns();

                xlPackage.Save();
                logger.Info("Completed");
            }
        }

        public StringBuilder ConvertDataTableToCsvFile(DataTable dtData)
        {
            StringBuilder data = new StringBuilder();

            //Taking the column names.
            for (int column = 0; column < dtData.Columns.Count; column++)
            {
                //Making sure that end of the line, shoould not have comma delimiter.
                if (column == dtData.Columns.Count - 1)
                    data.Append(dtData.Columns[column].ColumnName.ToString().Replace(",", ";"));
                else
                    data.Append(dtData.Columns[column].ColumnName.ToString().Replace(",", ";") + ',');
            }

            data.Append(Environment.NewLine);//New line after appending columns.

            for (int row = 0; row < dtData.Rows.Count; row++)
            {
                for (int column = 0; column < dtData.Columns.Count; column++)
                {
                    ////Making sure that end of the line, shoould not have comma delimiter.
                    if (column == dtData.Columns.Count - 1)
                        data.Append(dtData.Rows[row][column].ToString().Replace(",", ";"));
                    else
                        data.Append(dtData.Rows[row][column].ToString().Replace(",", ";") + ',');
                }

                //Making sure that end of the file, should not have a new line.
                if (row != dtData.Rows.Count - 1)
                    data.Append(Environment.NewLine);
            }
            return data;
        }

        //This method saves the data to the csv file. 
        public void SaveData(StringBuilder data, string filePath)
        {
            using (StreamWriter objWriter = new StreamWriter(filePath))
            {
                objWriter.WriteLine(data);
            }
        }

    }
}
