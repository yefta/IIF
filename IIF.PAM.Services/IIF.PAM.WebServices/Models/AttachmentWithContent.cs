using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml.Linq;

namespace IIF.PAM.WebServices.Models
{
    public class AttachmentWithContent
    {
        private AttachmentTypeConstants _AttachmentType;
        public AttachmentTypeConstants AttachmentType
        {
            get
            {
                return this._AttachmentType;
            }
            set
            {                
                this._AttachmentType = value;
                this.AttachmentTypeDisplayName = this._AttachmentType.ToDisplayName();
                this.AttachmentTypeDMSMetadataDisplayName = this._AttachmentType.ToDMSMetadataDisplayName();
            }
        }
        public string AttachmentTypeDisplayName { get; private set; }
        public string AttachmentTypeDMSMetadataDisplayName { get; private set; }        
        public string Attachment { get; set; }
        public string FileName { get; private set; }
        public byte[] FileContent { get; private set; }
        public XDocument Document { get; private set; }
        public string Description { get; set; }

        public void ParseAttachment()
        {
            this.Document = XDocument.Parse(this.Attachment);
            this.FileName = this.Document.Root.Element("name").Value;
            if (this.FileName.ToLower() == "scnull")
            {
                this.FileName = null;
            }
            if (this.Document.Root.Element("content").Value.ToLower() == "scnull")
            {
                this.FileContent = null;
            }
            else
            {
                this.FileContent = Convert.FromBase64String(this.Document.Root.Element("content").Value);
            }
        }
    }
}