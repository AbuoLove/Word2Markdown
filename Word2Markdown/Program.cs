using System.Configuration;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Word2MarkDown
{
    class Program
    {
        static void Main(string[] args)
        {
            var inDir = ConfigurationManager.AppSettings.Get("inDir");
            var outDir = ConfigurationManager.AppSettings.Get("outDir");

            string[] names = Directory.GetFiles(inDir, "*.docx");
            foreach (string file in names)
            {
                StringBuilder sb = new StringBuilder();
                string fileName = Path.GetFileName(file);
                string mdFile = outDir + @"\" + fileName.Replace(".docx",".md");
                using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(file, true))
                {
                    Body body = wdDoc.MainDocumentPart.Document.Body;
                    var bodyChildren = body.Descendants<Paragraph>();
                    

                    foreach (Paragraph p in bodyChildren)
                    {
                        string header = "\n";
                        string st = string.Empty;

                        if (p.ParagraphProperties != null && p.ParagraphProperties.ParagraphStyleId != null)
                        {
                            st= p.ParagraphProperties.ParagraphStyleId.Val;
                        }

                        if (!string.IsNullOrEmpty(st))
                        {
                            header = convHeader(st);
                        }

                        sb.Append(header + p.InnerText);
                    }

                }

                using (var sw = new StreamWriter(mdFile, false))
                {
                    sw.Write(sb.ToString());
                }
            }
        }

        static string convHeader(string pstyle)
        {
            string retStr = string.Empty;

            switch (pstyle)
            {
                case "Heading1":
                    retStr = "\n# ";
                    break;
                case "Heading2":
                    retStr = "\n## ";
                    break;
                case "Heading3":
                    retStr = "\n### ";
                    break;
                case "Heading4":
                    retStr = "\n#### ";
                    break;
                case "Heading5":
                    retStr = "\n##### ";
                    break;
                default:
                    retStr = "\n\n";
                    break;
            }

            return retStr;
        }
    }
}
