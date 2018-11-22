using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace IIF.PAM.MergeDocumentServices.Models
{
    public class FileMergeResult
    {
        public byte[] FileContent { get; set; }
        public string FileName { get; set; }
    }
}
