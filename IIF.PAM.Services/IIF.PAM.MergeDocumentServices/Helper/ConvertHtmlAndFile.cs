using System;
using System.IO;
using System.Text.RegularExpressions;
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
                html = string.Format("<html style=\"font-family: Roboto Light; font-size: 13px\">{0}</html>", html);
                writer.WriteLine(html);
            }
            return htmlTempFilePath;
        }

		public static string SaveToHtmlNew(string html, string fontFamily, float fontSize)
		{
			string htmlTempFilePath = Path.Combine(Path.GetTempPath(), string.Format("{0}.html", Path.GetRandomFileName()));
			string myFontSize = convertFontSize(fontSize);

			//apus <style>
			html = Regex.Replace(html, "(<style.+?</style>)|(<script.+?</script>)", "", RegexOptions.IgnoreCase | RegexOptions.Singleline);
			//apus style dari s sampai >
			html = Regex.Replace(html, "(style.+?>)", ">", RegexOptions.IgnoreCase | RegexOptions.Singleline);

			using (StreamWriter writer = File.CreateText(htmlTempFilePath))
			{
				html = string.Format("<html style=\"font-family: "+ fontFamily + "; font-size: "+ myFontSize + "\">{0}</html>", html);
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

		public static string SaveToFileTmp(string sFile)
		{			
			XDocument xDoc = new XDocument();
			xDoc = XDocument.Parse(sFile);			
			byte[] bFile = Convert.FromBase64String(xDoc.Root.Element("content").Value);			
			
			var tmpFile = Path.GetTempFileName();
			File.WriteAllBytes(tmpFile, bFile);

			return tmpFile;
		}

		public static string convertFontSize(float fontSize)
		{
			string res = "";

			if (fontSize == 10)
				res = "13px";

			return res;
		}
    }
}