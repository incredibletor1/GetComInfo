using System.Diagnostics;

namespace GetComInfo
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Process[] vsProcs = Process.GetProcessesByName("cmd");
            if (vsProcs.Any())
            {
                var maxTime = vsProcs.Max(p => p.StartTime);
                vsProcs.First(p => p.StartTime == maxTime).Kill();
            }

            ApplicationConfiguration.Initialize();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}