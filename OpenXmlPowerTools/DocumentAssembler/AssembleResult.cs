using System.Collections.Generic;

namespace OpenXmlPowerTools
{
    public partial class DocumentAssembler
    {
        internal class AssembleResult
        {
            public bool HasError;
            public List<AssembleInsert> Inserts = new List<AssembleInsert>();
        }
    }
}