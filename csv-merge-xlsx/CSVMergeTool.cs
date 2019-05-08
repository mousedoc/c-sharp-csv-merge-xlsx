using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace csv_merge_xlsx
{
    public static class CSVMergeToolDefine
    {
        public static readonly string ExcelFileName = "data";

        public static string ExcelFileNameWithExt
        {
            get { return ExcelFileName + ".xlsx"; }
        }
    }

    class CSVMergeTool
    {
        private Microsoft.Office.Interop.Excel.Application app = null;
        private object missValue = System.Reflection.Missing.Value;

        public object Excel { get; private set; }

        public void Run()
        {
            Workbook targetBook = null;

            try
            {
                string dataFileFullPath= Path.Combine(Directory.GetCurrentDirectory(), CSVMergeToolDefine.ExcelFileName);
                string dataFileFullPathNoExt = Path.Combine(Directory.GetCurrentDirectory(), CSVMergeToolDefine.ExcelFileName);

                if (File.Exists(dataFileFullPath))
                {
                    File.Delete(dataFileFullPath);
                }

                this.app = new Microsoft.Office.Interop.Excel.Application();
                this.app.DisplayAlerts = false;

                if (app == null)
                {
                    Logger.Error("Excel is not properly installed");
                    return;
                }

                var csvFiles = GetCSVFiles();
                if (csvFiles.Count < 0)
                {
                    Logger.Error("Have no .csv files");
                    return;
                }

                targetBook = app.Workbooks.Add(missValue);

                // Copy sheets
                foreach (var path in csvFiles)
                    CopyTargetCSVSheet(path, targetBook);

                Logger.Warn("Removing empty sheet...");
                Worksheet emptySheet = targetBook.Worksheets["Sheet1"];
                emptySheet.Delete();

                Logger.Warn("Save data.xlsx...");
                targetBook.Activate();
                targetBook.SaveAs(
                    dataFileFullPathNoExt,
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,
                    Missing.Value,
                    Missing.Value,
                    false,
                    false,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution,
                    true,
                    Missing.Value, Missing.Value, Missing.Value);

                Logger.Log("Success");
            }
            catch (Exception exception)
            {
                Logger.Error(exception.ToString());
                Logger.Error("Fail");
            }
            finally
            {
                Release();
            }
        }

        public void Release()
        {
            if (app == null)
                return;

            foreach (Workbook book in app.Workbooks)
                book.Close();

            if (app != null)
                app.Quit();

            Marshal.FinalReleaseComObject(app);
            app = null;
        }

        private List<string> GetCSVFiles()
        {
            var allFiles = Directory.GetFiles("./");
            var csvFiles = new List<string>();

            Logger.Log("List of .csv");
            foreach (var localpath in allFiles)
            {
                if (localpath.EndsWith(".csv") == false)
                    continue;

                string fullPath = Path.GetFullPath(localpath);
                csvFiles.Add(fullPath);
                Logger.Log("\t" + Path.GetFileName(fullPath));
            }
            Logger.Log("\n\n");
            return csvFiles;
        }

        private Microsoft.Office.Interop.Excel.Worksheet CopyTargetCSVSheet(string path, Microsoft.Office.Interop.Excel.Workbook workbook)
        {
            var sourceBook = app.Workbooks.Open(path);

            string name = Path.GetFileNameWithoutExtension(path);
            var sheet = (Microsoft.Office.Interop.Excel.Worksheet)app.Worksheets[name];
            if (sheet == null)
            {
                Logger.Error("GetTargetCSVSheet(...) - sheet is null");
                return null;
            }

            string lastSheetName = string.Empty;
            foreach (Worksheet workSheet in workbook.Worksheets)
            {
                lastSheetName = workSheet.Name;
            }

            Logger.Warn(string.Format("Copy {0}...", name));
            sheet.Copy(workbook.Worksheets[lastSheetName]);
            sourceBook.Close();
            return sheet;
        }
    }
}
