using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text;

using IIF.PAM.MergeDocumentServices;

namespace IIF.PAM.MergeDocumentTester
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
		        log4net.Config.XmlConfigurator.Configure();
                MergeDocument svcMerge = new MergeDocument();
				//string conStringIIF = "data source=k2projectiif;initial catalog=IIF;user id=sa;password=P@ssw0rd;";
				//string conStringIIF = "data source=.;initial catalog=IIF;user id=sa;password=P@ssw0rd;";
				//string templateFolder = @"\\k2projectiif\c$\IIF\PAMTemplate\Newest";
				//string mergeResultFolder = @"\\k2projectiif\c$\IIF\MergeResult";

				//svcMerge.MergePAMDocument(32, conStringIIF, @"D:\Srf\Project\PIS\IIF\Merge\PAMTemplate", @"D:\Srf\Project\PIS\IIF\Merge\Temp", "MergeByFQN", "MergeBy");
				//svcMerge.MergePAMDocument(10112, conStringIIF, templateFolder, @"C:\temp", "MergeByFQN", "MergeBy");

				#region DB Old (IIF_Report)
				//10168 - CM Project Finance
				//svcMerge.MergeCMDocument(10168, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//20210 - CM Corporate Finance
				//svcMerge.MergeCMDocument(20210, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//53 - CM Equity
				//svcMerge.MergeCMDocument(53, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//20237 - CM Waiver
				//svcMerge.MergeCMDocument(20237, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");


				//10131 - PAM Project Finance (gede)
				//svcMerge.MergePAMDocument(10131, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");
				//10121 - PAM Project Finance (standard)
				//svcMerge.MergePAMDocument(10121, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");
				//svcMerge.MergePAMDocument(10121, conStringIIF, templateFolder, @"C:\temp", "MergeByFQN", "MergeBy");

				//10112 - PAM Corporate Finance
				//svcMerge.MergePAMDocument(10112, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");
				//82 - PAM Corporate Finance (data aneh)
				//svcMerge.MergePAMDocument(82, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//10119 - PAM Equity
				//svcMerge.MergePAMDocument(10119, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");
				#endregion

				#region DBBARU-CLEAN (IIF)
				//string conStringIIF = "data source=.;initial catalog=IIF;user id=sa;password=P@ssw0rd;";
				//string templateFolder = @"C:\Project\IIF\IIF\PAMTemplate";
				//string mergeResultFolder = @"C:\Project\IIF\IIF\MergeResult";

				//42 - PAM Equity
				//svcMerge.MergePAMDocument(42, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//39 - PAM Equity
				//svcMerge.MergePAMDocument(39, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//75 - PAM Equity
				//svcMerge.MergePAMDocument(75, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//40 - PAM Corporate
				//svcMerge.MergePAMDocument(40, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//72 - PAM Corporate
				//svcMerge.MergePAMDocument(72, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//77 - PAM Corporate
				//svcMerge.MergePAMDocument(77, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//41 - PAM Project Finance
				//svcMerge.MergePAMDocument(41, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//70 - PAM Project Finance
				//svcMerge.MergePAMDocument(70, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//76 - PAM Project Finance
				//svcMerge.MergePAMDocument(76, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");




				//20 - CM Equity
				//svcMerge.MergeCMDocument(20, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//106 - CM Equity
				//svcMerge.MergeCMDocument(106, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//5 - CM Corporate Finance
				//svcMerge.MergeCMDocument(5, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");				

				//110 - CM Corporate Finance
				//svcMerge.MergeCMDocument(110, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");				

				//17 - CM Project Finance
				//svcMerge.MergeCMDocument(17, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//41 - CM Project Finance
				//svcMerge.MergeCMDocument(41, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//105 - CM Project Finance
				//svcMerge.MergeCMDocument(105, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//1 - CM Waiver
				//svcMerge.MergeCMDocument(1, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//103 - CM Waiver
				//svcMerge.MergeCMDocument(103, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");
				#endregion

				#region DB Dev/Prod IIF
				string conStringIIF = "data source=10.15.3.216\\IDJKTINSKTWO;initial catalog=IIF;user id=sa;password=@dmin@IIF.12;";
				string templateFolder = @"\\10.15.3.214\c$\IIF\PAMTemplate";
				string mergeResultFolder = @"\\10.15.3.214\c$\IIF\MergeResult";

				#region get INSTALLED fonts
				//InstalledFontCollection installedFontCollection = new InstalledFontCollection();				
				//FontFamily[] fontFamilies = installedFontCollection.Families;
				//int count = fontFamilies.Length;
				//for (int j = 0; j < count; ++j)
				//{
				//	Console.WriteLine(fontFamilies[j].Name);					
				//}
				#endregion

				//2 - PAM Project Finance
				//svcMerge.MergePAMDocument(2, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//10 - PAM Project Finance
				//svcMerge.MergePAMDocument(10, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//17 - PAM Project Finance
				//svcMerge.MergePAMDocument(17, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//26 - PAM Project Finance
				//svcMerge.MergePAMDocument(26, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//16 - PAM Equity
				//svcMerge.MergePAMDocument(16, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");


				//15 - CM WAIVER
				//svcMerge.MergeCMDocument(15, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//21 - CM PROJECT
				//svcMerge.MergeCMDocument(21, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//28 - CM PROJECT
				svcMerge.MergeCMDocument(28, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//23 - CM Corporate
				//svcMerge.MergeCMDocument(23, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//2 - CM Equity
				//svcMerge.MergeCMDocument(2, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//25 - CM Equity
				//svcMerge.MergeCMDocument(25, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");
				#endregion

				Console.WriteLine("Success");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.ReadLine();
        }
    }
}

