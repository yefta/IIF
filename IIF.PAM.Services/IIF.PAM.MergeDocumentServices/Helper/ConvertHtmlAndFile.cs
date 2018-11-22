using System;
using System.IO;
using System.Xml.Linq;

namespace IIF.PAM.MergeDocumentServices.Helper
{
    public class ConvertHtmlAndFile
    {
        public static string SaveToHtml(string html)
        {
            string htmlTempFilePath = Path.Combine(Path.GetTempPath(), string.Format("{0}.html", Path.GetRandomFileName()));
            using (StreamWriter writer = File.CreateText(htmlTempFilePath))
            {
                html = string.Format("<html>{0}</html>", html);
                writer.WriteLine(html);
            }
            return htmlTempFilePath;
        }

        public static string SaveToFile(string sFile)
        {
            String ID = Guid.NewGuid().ToString().ToUpper();
            XDocument xDoc = new XDocument();
            xDoc = XDocument.Parse(sFile);
            string sFileName = xDoc.Root.Element("name").Value;
            string extensionFile = Path.GetExtension(sFileName);
            string fileName = String.Format("{0}{1}", ID, extensionFile);
            byte[] bFile = Convert.FromBase64String(xDoc.Root.Element("content").Value);
            var tmpFile = Path.GetTempFileName();
            File.WriteAllBytes(tmpFile, bFile);

            return tmpFile;
        } 
    }
}