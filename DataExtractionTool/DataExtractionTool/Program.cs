using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using System.IO;
using System.Data.SqlClient;
using System.Data;


namespace DataExtractionTool
{
    class Program
    {
        public static string ProjectFolder;
       
        static void Main(string[] args)
        {
            Stopwatch sw = new Stopwatch();
            ProjectFolder = Directory.GetCurrentDirectory();
 
            SqlConnection sqlConnection = new SqlConnection();

            List<DataCategory> dataCategories = new List<DataCategory>();
            OutputSettings outputSettings = new OutputSettings();
            List<OutColumn> columns = new List<OutColumn>();
            string Inpfile = String.Concat(ProjectFolder, "/input.xlsx");
            string inptColum = "InputValue";
            string tmpTableName = "#DataExtractTool";

            Excel.Application xl = new Excel.Application();
            sw.Start();
            var LastRecord = sw.Elapsed;

            try
            {

                if (!File.Exists(Inpfile))
                {
                    throw new Exception($"File {Inpfile} does not exist");
                }


                Excel.Workbook wb = xl.Workbooks.Open(Inpfile);
                GetInputSettings(dataCategories,outputSettings,columns, wb, "Inputs");
                LastRecord = PrintTime(sw,LastRecord);
                //*****CreateSqlConnection
                sqlConnection = CreateConnection("PSRVWDB1663\\PSQLPBI0002",dataCategories[1].SrcDB);
                LastRecord = PrintTime(sw, LastRecord);
                ReadInputDataSendToSQLServer(wb,"InputData", sqlConnection, tmpTableName, inptColum, 5000);
                LastRecord = PrintTime(sw, LastRecord);
                //**Interpret and apply OutputSettings

                //******execute SQL Query --All SQL logic is contained in below function
                ExecuteSQL(sqlConnection, dataCategories, outputSettings, columns,tmpTableName,inptColum);
                LastRecord = PrintTime(sw, LastRecord);



                wb.Close(true, null, null);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wb);
                wb = null;

                xl.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xl);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xl);

                xl = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();



            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                Logging.WriteErrorToFile(ProjectFolder, e);


                xl.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xl);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xl);
                xl = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();



            }

        }

        //**************All logic regarding the settings will happen here*********************
        public static void ExecuteSQL(SqlConnection conn, List<DataCategory> dataCategories, OutputSettings outputSettings, List<OutColumn> columns, string tmpTableName, string inptColum)
        {

            string SQL = "SELECT ";
            int cnt = 1;
            foreach (OutColumn c in columns)
            {
                if (c.Use)
                {
                    if (cnt == 1)
                    {
                        SQL = SQL + c.Short + "." + c.SelColumn;
                    }
                    else
                    {
                        SQL = SQL + Environment.NewLine + "," + c.Short + "." + c.SelColumn;
                    }

                }
                cnt += 1;
            }

            Console.WriteLine(SQL);

            //SqlCommand cmd = new SqlCommand("SELECT COUNT(0) cnt FROM " + tmpTableName,conn);
            //using (SqlDataReader reader = cmd.ExecuteReader())
            //{
            //    while (reader.Read())
            //    {
            //        Console.WriteLine(reader["cnt"].ToString());
            //    }
            //}
        }
        
        public static System.TimeSpan PrintTime(Stopwatch sw, System.TimeSpan LastRecord)
        {
            Console.WriteLine("Time spent on last process : " + Convert.ToString(sw.Elapsed - LastRecord));
            Console.WriteLine("Time spent total : " + sw.Elapsed);
            LastRecord = sw.Elapsed;

            return LastRecord;
        } 
        public static void ReadInputDataSendToSQLServer(Excel.Workbook wb,string sheetStr, SqlConnection conn, string tmpTableName, string inptColum,int BulkSize = 10)
        {

            Console.WriteLine("--------Inserting rows to temporary table on SQL Server--------");

            SqlCommand cmd = new SqlCommand();

            cmd.Connection = conn;
            cmd.CommandText = "DROP TABLE IF EXISTS "+tmpTableName+";"
                +" CREATE TABLE "+tmpTableName+" ("+ inptColum + " varchar(1000))";
            cmd.ExecuteNonQuery();


            Excel.Worksheet ws = wb.Worksheets[sheetStr];
            SqlBulkCopy bulkCpy = new SqlBulkCopy(conn);
            bulkCpy.DestinationTableName = tmpTableName;
            bulkCpy.ColumnMappings.Add(inptColum, inptColum);


            DataTable tbl = new DataTable();
            tbl.Columns.Add("InputValue");
            double i = 2;
            while ((string)(ws.Cells[i, 1] as Excel.Range).Value!=null )
            {
                tbl.Rows.Add((string)(ws.Cells[i, 1] as Excel.Range).Value);
                
                if (i % BulkSize == 0)
                {
                    bulkCpy.WriteToServer(tbl);
                    tbl.Rows.Clear(); //Reset rows so we are not overloading the ram
                    Console.WriteLine(" Rows Inserted : " + tbl.Rows.Count.ToString());
                }

                i += 1;
            }
            
            if (tbl.Rows.Count > 0)
            {
                bulkCpy.WriteToServer(tbl);
                Console.WriteLine("Rows Inserted : "+ tbl.Rows.Count.ToString());
                tbl.Rows.Clear();
            }

            
            
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ws);
        }

        public static SqlConnection CreateConnection(string Server,string InitialDatabase)
        {
            Console.WriteLine("--------Creating connection to SQL Server--------");
            SqlConnection conn = new SqlConnection();
            string connString = "Data Source=PSRVWDB1663\\PSQLPBI0002;Initial Catalog="+ InitialDatabase +";Integrated Security=true;";
            conn.ConnectionString = connString;
            conn.Open();

            return conn;
        }



        public static void GetInputSettings(List<DataCategory> dataCategories, OutputSettings outputSettings, List<OutColumn> columns,Excel.Workbook wb,string InptSheetName)
        {
            

            int i = 1;
            string currInpCat = null;
            
            Excel.Worksheet sheet = wb.Worksheets[InptSheetName];
            Console.WriteLine("--------Reading input settings from : " + sheet.Name + "--------");



            string col1;
            string col2;

            while ((string)(sheet.Cells[i, 1] as Excel.Range).Value != "EndOfInput")
            {
                col1 = (string)(sheet.Cells[i, 1] as Excel.Range).Value;
                //Data Category

                if (col1 == "Data Category")
                {
                    currInpCat = "Data Category";
                }
                else if (col1 == "Output Settings")
                {
                    currInpCat = "Output Settings";
                }
                else if (col1 == "Columns")
                {
                    currInpCat = "Columns";
                }

                if (col1 != currInpCat && col1 != null)
                {
                    if (currInpCat == "Data Category")
                    {
                        DataCategory dc = new DataCategory();

                        dc.Category = (string)(sheet.Cells[i, 1] as Excel.Range).Value;
                        dc.Use = (bool)(sheet.Cells[i, 2] as Excel.Range).Value;
                        dc.SrcDB = (string)(sheet.Cells[i, 3] as Excel.Range).Value;
                        dc.SrcSchema = (string)(sheet.Cells[i, 41] as Excel.Range).Value;
                        dc.SrcTable = (string)(sheet.Cells[i, 5] as Excel.Range).Value;
                        dc.Short = (string)(sheet.Cells[i, 6] as Excel.Range).Value;

                        dataCategories.Add(dc);
                    }

                    if (currInpCat == "Output Settings")
                    {
                
                        string val = (string)(sheet.Cells[i, 1] as Excel.Range).Value;
                        switch (col1)
                        {
                            case "Output Type":
                                outputSettings.OutputType = val;
                                break;
                            case "Aggregation level":
                                outputSettings.AggregationLevel = val;
                                break;
                            case "From":
                                outputSettings.From = Convert.ToDateTime(val);
                                break;
                            case "To":
                                outputSettings.To = Convert.ToDateTime(val);
                                break;
                                

                        }


                    }


                        //}
                        if (currInpCat == "Columns")
                    {
                        OutColumn o = new OutColumn();
                        o.SelColumn = (string)(sheet.Cells[i, 1] as Excel.Range).Value;
                        o.Use = (bool)(sheet.Cells[i, 2] as Excel.Range).Value;
                        o.DataCategory = (string)(sheet.Cells[i, 3] as Excel.Range).Value;
                        o.Short = (string)(sheet.Cells[i, 4] as Excel.Range).Value;

                        columns.Add(o);
                    }

                }




                i += 1;
            }




            //sheet.SaveAs(string.Concat(ProjectFolder,"/input.csv"),Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV,);


            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sheet);
            sheet = null;

            

            
            //return ReturnArray;
        }
    }

    public class DataCategory
    {
        public string Category { get; set; }
        public bool Use { get; set; }
        public string SrcDB { get; set; }
        public string SrcSchema { get; set; }
        public string SrcTable { get; set; }
        public string Short { get; set; }
    }
    public class OutputSettings
    {
        public string OutputType { get; set; }
        public string AggregationLevel { get; set; }
        public DateTime From { get; set; }
        public DateTime To { get; set; }

    }
    public class OutColumn
    {
        public string SelColumn { get; set; }
        public bool Use { get; set; }
        public string DataCategory { get; set; }
        public string Short { get; set; }


    }
 
}
