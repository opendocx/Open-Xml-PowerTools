namespace OpenXmlPowerTools
{
    public class DocxSource : DocumentBuilder.Source
    {
        public DocxSource(string fileName)
            : base(fileName)
        {
        }

        public DocxSource(WmlDocument source)
            : base(source)
        {
        }

        public DocxSource(string fileName, bool keepSections)
            : base(fileName, keepSections)
        {
        }

        public DocxSource(WmlDocument source, bool keepSections)
            : base(source, keepSections)
        {
        }

        public DocxSource(string fileName, string insertId)
            : base(fileName, insertId)
        {
        }

        public DocxSource(WmlDocument source, string insertId)
            : base(source, insertId)
        {
        }
    }
}
