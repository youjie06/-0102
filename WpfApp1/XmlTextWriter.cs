using System.IO;

namespace _2023_WpfApp4
{
    internal class XmlTextWriter
    {
        private StringWriter stringWriter;

        public XmlTextWriter(StringWriter stringWriter)
        {
            this.stringWriter = stringWriter;
        }
    }
}