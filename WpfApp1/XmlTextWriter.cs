using System.IO;

namespace _2023_WpfApp4
{
    internal class XmlTextWriter
    {
        private StringWriter stringWriter;

        // 初始化 XmlTextWriter 類別的新實例，用來接收 XML 數據的 StringWriter
        public XmlTextWriter(StringWriter stringWriter)
        {
            this.stringWriter = stringWriter;
        }
    }
}