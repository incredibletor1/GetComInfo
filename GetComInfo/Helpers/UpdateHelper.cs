using System.Diagnostics;

namespace GetComInfo.Helpers
{
    /// <summary>
    /// Update Helper class
    /// </summary>
    public static class UpdateHelper
    {
        /// <summary>
        /// Build Cmd process
        /// </summary>
        public static void Cmd(string line)
        {
            Process.Start(new ProcessStartInfo { FileName = "cmd", WorkingDirectory = Environment.CurrentDirectory, Arguments = $"/c {line}", WindowStyle = ProcessWindowStyle.Hidden });
        }
    }
}
