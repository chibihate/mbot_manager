using System;
using System.IO;
using System.Net;
using System.Text;

namespace MBotServerLocally
{
    public static class WebServer
    {
        public static void StartServer()
        {
            WebServer.webServer = new HttpListener();
            WebServer.webServer.Prefixes.Add("http://127.0.0.1:8055/");
            WebServer.webServer.Start();
            WebServer.Accept();
        }

        private static void Accept()
        {
            try
            {
                WebServer.webServer.BeginGetContext(new AsyncCallback(WebServer.Accept2), null);
            }
            catch
            {
            }
        }

        private static void ProcessRequest(HttpListenerContext context)
        {
            try
            {
                HttpListenerRequest request = context.Request;
                bool flag = request.Url.ToString().Contains("pa/auth.psro.mbot.1.php");
                if (flag)
                {
                    HttpListenerResponse response = context.Response;
                    string text = "@0.2CEA9DD0D41A2E06FD2F3D628566247BB14E114E594767090FEDAFC3169D92CA2CA3918B2EB4C38E3D8F38822D9754E0238C3DFD4DFF50945096248B398330842BEF2BEF2B84408440842BEF2BEF2B5419A80A09CA1D07.AC69C053AA000000009874780E@";
                    response.ContentLength64 = (long)text.Length;
                    Stream outputStream = response.OutputStream;
                    outputStream.Write(Encoding.UTF8.GetBytes(text), 0, text.Length);
                    outputStream.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        internal static void Accept2(IAsyncResult iar)
		{
			HttpListenerContext context = WebServer.webServer.EndGetContext(iar);
            WebServer.ProcessRequest(context);
			WebServer.Accept();
		}

        private static HttpListener webServer = null;
    }
}
