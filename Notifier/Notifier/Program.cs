using System.IO;
using System.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Notifier
{
    class Program
    {

        public static string FolderName { get; set; }
        public static string FolderPath { get; set; }
        public static string KeywordsFileName { get; set; }
        public static string ExcelFileName { get; set; }
        public static string vbsfileName { get; set; }
        public static List<string> KeywordsSelected { get; set; }
        public static string ColumnName { get; set; }

        static void Main(string[] args)
        {
            getconfigDetails();
            FileWatch();
            while (true) ;
        }

        static DataRowCollection LoadWorksheetInDataTable(string fileName = null, string sheetName = null)
        {
            System.Data.DataTable sheetData = new System.Data.DataTable();
            DataRowCollection results = null;
            using (System.Data.OleDb.OleDbConnection conn = returnConnection(fileName))
            {
                conn.Open();
                System.Data.OleDb.OleDbDataAdapter sheetAdapter = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", conn);
                sheetAdapter.Fill(sheetData);
                results = sheetData.Clone().Rows;
            }

            foreach (DataRow dr in sheetData.Rows)
            {
                foreach (string item in KeywordsSelected)
                {
                    if (dr.Field<string>(ColumnName).Trim().Contains(item))
                    {
                        results.Add(dr.ItemArray);
                    }
                }


            }

            return results;
        }

        private static System.Data.OleDb.OleDbConnection returnConnection(string fileName)
        {
            return new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path.Combine(FolderName, ExcelFileName) + "; Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\"");
        }
        private static void ReadXls()
        {
            var y = LoadWorksheetInDataTable();
            if (y.Count > 0)
            {
                System.Diagnostics.Process scriptProc = new System.Diagnostics.Process();
                scriptProc.StartInfo.FileName = Path.Combine(FolderName, vbsfileName);
                scriptProc.Start();
                scriptProc.WaitForExit();
                scriptProc.Close();
            }
        }

        private static void FileWatch()
        {
            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = FolderName;
            watcher.Filter = ExcelFileName;
            watcher.Created += new FileSystemEventHandler(Watcher_Created);
            watcher.EnableRaisingEvents = true;
        }

        private static List<string> readKeywords()
        {
            return File.ReadLines(Path.Combine(FolderName, KeywordsFileName)).ToList<string>();
        }

        private static void getconfigDetails()
        {
            FolderName = ConfigurationSettings.AppSettings.GetValues("FoldertoWaatch").SingleOrDefault().ToString();
            KeywordsFileName = ConfigurationSettings.AppSettings.GetValues("KeywordsFileName").SingleOrDefault().ToString();
            ExcelFileName = ConfigurationSettings.AppSettings.GetValues("ExcelFileName").SingleOrDefault().ToString();
            vbsfileName = ConfigurationSettings.AppSettings.GetValues("vbsfileName").SingleOrDefault().ToString();
            ColumnName = ConfigurationSettings.AppSettings.GetValues("ColumnName").SingleOrDefault().ToString();
        }

        private static void Watcher_Created(object sender, FileSystemEventArgs e)
        {

            KeywordsSelected = readKeywords();
            ReadXls();
            Console.WriteLine("File Created");
        }
    }
}