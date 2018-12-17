using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IIF.PAM.WebServices
{
    public class ApplicationConfig
    {
        public string IIFConnectionString { get; set; }
        public string WebServiceUrl { get; set; }
        public string ReportViewerUrl { get; set; }
        public string RichTextEditorUrl { get; set; }
        public string K2Server { get; set; }
        public string PAMMergeDocumentTemplateLocation { get; set; }
        public string PAMMergeDocumentTemporaryLocation { get; set; }
        public string CMMergeDocumentTemplateLocation { get; set; }
        public string CMMergeDocumentTemporaryLocation { get; set; }
        public string SmartObjectName_ADUser { get; set; }
        public string WorkflowNames { get; set; }

        public string SMTPFromEmail { get; set; }
        public string SMTPFromName { get; set; }
        public string SMTPHost { get; set; }
        public int SMTPPort { get; set; }
        public bool SMTPEnableSSL { get; set; }
        public string SMTPCredentialName { get; set; }
        public string SMTPCredentialPassword { get; set; }
    }
}