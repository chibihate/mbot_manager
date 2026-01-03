using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace KillMbotServerLocally
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Program.HideConsole();
            foreach (Process process in Process.GetProcessesByName("mBotServerLocally"))
            {
                try
                {
                    process.Kill();
                    process.WaitForExit();
                }
                catch
                {
                }
            }
        }

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        private static void HideConsole()
        {
            Program.ShowWindow(Program.GetConsoleWindow(), 0);
        }

        private const int SW_HIDE = 0;
    }
}
