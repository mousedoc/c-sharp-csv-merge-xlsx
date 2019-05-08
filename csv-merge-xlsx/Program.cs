using System;

namespace csv_merge_xlsx
{
    internal class Program
    {
        private static CSVMergeTool tool = null;

        static void Main(string[] args)
        {
            tool = new CSVMergeTool();
            tool.Run();

            Logger.Warn("\r\n\nPress a key");
            Console.ReadKey();
        }

        static void OnProcessExit(object sender, EventArgs args)
        {
            if (tool == null)
                return;

            tool.Release();
        }
    }
}
