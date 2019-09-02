using System;

namespace csv_merge_xlsx
{
    internal class Program
    {
        public static readonly float Version = 2.0f;

        private static CSVMergeTool tool = null;

        private static void Main(string[] args)
        {
            ShowInfo();

            Run();

            PressAKey();
        }

        private static void ShowInfo()
        {
            Logger.Warn("Info");
            Logger.Log("    Version - " + Version.ToString("0.0"));
            Logger.Log("");
        }

        private static void Run()
        {
            tool = new CSVMergeTool();
            tool.Run();
        }

        private static void PressAKey()
        {
            Logger.Warn("\r\n\nPress a key");
            Console.ReadKey();
        }

        private static void OnProcessExit(object sender, EventArgs args)
        {
            if (tool == null)
                return;

            tool.Release();
        }
    }
}
