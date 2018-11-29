﻿using System;
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
				//svcMerge.MergePAMDocument(32, conStringIIF, @"D:\Srf\Project\PIS\IIF\Merge\PAMTemplate", @"D:\Srf\Project\PIS\IIF\Merge\Temp", "MergeByFQN", "MergeBy");
				//svcMerge.MergePAMDocument(10112, conStringIIF, @"C:\Project\IIF\PAMTemplate", @"C:\temp", "MergeByFQN", "MergeBy");
				svcMerge.MergeCMDocument(10191, conStringIIF, @"C:\Project\IIF\PAMTemplate", @"C:\temp", "MergeByFQN", "MergeBy");
				Console.WriteLine("Success");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //Console.ReadLine();
        }
    }
}
