using System.Collections.Generic;

namespace OpenXmlPowerTools
{
    public partial class WmlDocument : OpenXmlPowerToolsDocument
    {
        public IEnumerable<WmlDocument> SplitOnSections() => DocumentBuilder.DocumentBuilder.SplitOnSections(this);
    }
}