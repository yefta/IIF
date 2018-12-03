using System;
using System.Collections.Generic;
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
                string conStringIIF = "data source=k2projectiif;initial catalog=IIF;user id=sa;password=P@ssw0rd;";
				string templateFolder = @"\\k2projectiif\c$\IIF\PAMTemplate";
				string mergeResultFolder = @"\\k2projectiif\c$\IIF\MergeResult";
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
				//1 - PAM Equity
				//svcMerge.MergePAMDocument(1, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");								

				//2 - PAM Corporate
				//svcMerge.MergePAMDocument(2, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//5 - PAM Project Finance
				//svcMerge.MergePAMDocument(5, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");

				//1 - CM Waiver
				svcMerge.MergeCMDocument(1, conStringIIF, templateFolder, mergeResultFolder, "MergeByFQN", "MergeBy");
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

