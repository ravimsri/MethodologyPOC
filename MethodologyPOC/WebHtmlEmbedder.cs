using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Security;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace MethodologyPOC
{
    internal static class WebHtmlEmbedder
    {
        public static async Task InsertWebHtmlAtEndAsync(Word.Document doc, string url, int maxChars = 8000)
        {
            if (doc == null || string.IsNullOrWhiteSpace(url)) return;

            string html;
            try
            {
                html = await DownloadHtmlAsync(url);
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("InsertWebHtmlAtEndAsync" + ex.Message);
                return;
            }

            if (string.IsNullOrWhiteSpace(html)) return;

            if (html.Length > maxChars) html = html.Substring(0, maxChars);

            // cleanup html for word
            html = CleanHtmlForWord(html);

            // Wrap with title + URL
            html = WrapHtmlForWord(html, url);

            var tempPath = Path.Combine(Path.GetTempPath(), $"WordWebDemo_{Guid.NewGuid():N}.html");

            try
            {
                File.WriteAllText(tempPath, html);

                InsertHtmlFileAtEnd(doc, tempPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("HTML insert failed:\n" + ex.Message, "Web demo");
            }
            finally
            {
                try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
            }
        }

        private static async Task<string> DownloadHtmlAsync(string url)
        {
            //older .NET Framework defaults
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            using (var handler = new HttpClientHandler
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            })
            using (var http = new HttpClient(handler))
            {
                http.Timeout = TimeSpan.FromSeconds(30);
                http.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (WordAddinPOC)");
                return await http.GetStringAsync(url);
            }
        }

        private static string CleanHtmlForWord(string html)
        {
            // Remove scripts and styles for word
            html = Regex.Replace(html, "<script[\\s\\S]*?</script>", "", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, "<style[\\s\\S]*?</style>", "", RegexOptions.IgnoreCase);

            // Remove external stylesheet links for word
            html = Regex.Replace(html, "<link[\\s\\S]*?>", "", RegexOptions.IgnoreCase);

            // Remove navigation and layout for word
            html = Regex.Replace(html, "<nav[\\s\\S]*?</nav>", "", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, "<header[\\s\\S]*?</header>", "", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, "<footer[\\s\\S]*?</footer>", "", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, "<aside[\\s\\S]*?</aside>", "", RegexOptions.IgnoreCase);

            html = Regex.Replace(html, "<div[^>]*class=[\"']?mw-navigation[\"']?[\\s\\S]*?</div>", "", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, "<div[^>]*class=[\"']?vector-menu[\"']?[\\s\\S]*?</div>", "", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, "<div[^>]*class=[\"']?sidebar[\"']?[\\s\\S]*?</div>", "", RegexOptions.IgnoreCase);

            // Remove images and iframes for word
            html = Regex.Replace(html, "<img[^>]*>", "", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, "<iframe[\\s\\S]*?</iframe>", "", RegexOptions.IgnoreCase);

            // Remove noscript blocks for word
            html = Regex.Replace(html, "<noscript[\\s\\S]*?</noscript>", "", RegexOptions.IgnoreCase);

            // Remove skip links for word
            html = Regex.Replace(html, "<a[^>]*>Jump to content</a>", "", RegexOptions.IgnoreCase);

            return html;
        }


        private static string WrapHtmlForWord(string html, string url)
        {
            var safeUrl = SecurityElement.Escape(url);

            return $@"<!DOCTYPE html>
                <html>
                <head>
                <meta charset=""utf-8"">
                <title>Web Demo</title>
                <style>
                 body {{ font-family: Calibri, Arial; font-size: 11pt; }}
                 h1 {{ font-size: 16pt; }}
                 p  {{ line-height: 1.3; }}
                </style>
                </head>
                <body>
                <h1>DEMO OF - Microsoft Word’s capability to embed a web page within the document</h1>
                <p><b>Content from URL =</b> <a href=""{safeUrl}"">{safeUrl}</a></p>
                <hr/>
                {html}
                </body>
                </html>";
        }

        private static void InsertHtmlFileAtEnd(Word.Document doc, string htmlFilePath)
        {
            var end = doc.Content;
            end.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            end.InsertParagraphAfter();
            end.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            // Prevent the “Convert File” dialog
            doc.Application.Options.ConfirmConversions = false;

            end.InsertFile(
                FileName: htmlFilePath,
                Range: Type.Missing,
                ConfirmConversions: false,
                Link: false,
                Attachment: false
            );

            end.InsertParagraphAfter();
        }
    }
}
