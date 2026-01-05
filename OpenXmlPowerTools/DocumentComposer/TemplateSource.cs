using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class TemplateSource : DocxSource
    {
        public WmlDocument TemplateDoc { get; set; }
        public XElement Data { get; set; }
        public bool HasError { get; set; }

        public TemplateSource(string templateFileName, XmlDocument data)
            : this(templateFileName, data.GetXDocument().Root)
        { }

        public TemplateSource(string templateFileName, XElement data)
            : this(templateFileName, data, null)
        { }

        public TemplateSource(string templateFileName, XElement data, string insertId) : base((WmlDocument)null, insertId)
        {
            TemplateDoc = new WmlDocument(templateFileName);
            Data = data;
            HasError = false;
        }

        public async Task DoAssembly()
        {
            await Task.Yield();
            WmlDocument = DocumentAssembler.AssembleDocument(TemplateDoc, Data, out DocumentAssembler.AssembleResult results);
            HasError = results.HasError;
        }
    }
}
