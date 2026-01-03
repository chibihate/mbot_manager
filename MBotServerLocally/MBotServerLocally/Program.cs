using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace MBotServerLocally
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Program.HideConsole();
            WebServer.StartServer();
            new Thread(delegate ()
            {
                while (Console.ReadLine() != null)
                {
                    Thread.Sleep(100);
                }
            }).Start();
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
