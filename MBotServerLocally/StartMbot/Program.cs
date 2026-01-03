using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace StartMbot
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Program.HideConsole();
            if (Process.GetProcessesByName("mBotServerLocally").Length == 0)
            {
                Process.Start("mBotServerLocally.exe");
            }
            Process.Start("mBotLoader.exe");
            Thread.Sleep(5000);
            File.Delete("merrsend.exe");
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
