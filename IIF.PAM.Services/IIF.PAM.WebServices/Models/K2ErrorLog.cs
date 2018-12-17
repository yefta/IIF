using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using SourceCode.SmartObjects.Client;

namespace IIF.PAM.WebServices.Models
{
    public class K2ErrorLog
    {
        public int Id { get; set; }
        public int ProcID { get; set; }

        public int ProcInstID { get; set; }

        public int ObjectID { get; set; }
        public int TypeID { get; set; }

        public string Description { get; set; }

        public string ProcessName { get; set; }

        public K2ErrorLog (SmartObject soItem)
        {            
            this.Id = Convert.ToInt32(soItem.Properties["ID"].Value);
            this.ProcID = Convert.ToInt32(soItem.Properties["ProcID"].Value);
            this.ProcInstID = Convert.ToInt32(soItem.Properties["ProcInstID"].Value);
            this.ObjectID = Convert.ToInt32(soItem.Properties["ObjectID"].Value);
            this.TypeID = Convert.ToInt32(soItem.Properties["TypeID"].Value);
            this.Description = soItem.Properties["Description"].Value;
            this.ProcessName = soItem.Properties["ProcessName"].Value;
        }
    }
}