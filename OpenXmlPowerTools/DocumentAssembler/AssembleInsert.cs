using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public partial class DocumentAssembler
    {
        internal class AssembleInsert
        {
            public string Id { get; set; }
            public XElement Data { get; set; }
        }
    }
}
