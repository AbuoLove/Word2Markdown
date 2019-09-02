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
                string fileName = Path.GetFileName(file);
                string mdFile = outDir + @"\" + fileName.Replace(".docx",".md");
                string mdText = string.Empty;
                using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(file, true))
                {
                    Body body = wdDoc.MainDocumentPart.Document.Body;
                    mdText = convString(body);
                }

                using (var sw = new StreamWriter(mdFile, false))
                {
                    sw.Write(mdText);
                }
            }
        }

        static string convString(Body body)
        {
            StringBuilder sb = new StringBuilder();
            var bodyChildren = body.Descendants<Paragraph>();

            foreach (Paragraph p in bodyChildren)
            {
                string pstyle = string.Empty;
                string retStr = string.Empty;

                if (p.ParagraphProperties != null && p.ParagraphProperties.ParagraphStyleId != null)
                {
                    pstyle = p.ParagraphProperties.ParagraphStyleId.Val;
                }

                if (!string.IsNullOrEmpty(pstyle))
                {
                    switch (pstyle)
                    {
                        case "Title":
                            sb.Append("---");
                            sb.Append("title: " + p.InnerText);
                            sb.Append("---");
                            break;
                        case "Heading1":
                            sb.Append("\n# ");
                            sb.Append(p.InnerText);
                            break;
                        case "Heading2":
                            sb.Append("\n## ");
                            sb.Append(p.InnerText);
                            break;
                        case "Heading3":
                            sb.Append("\n### ");
                            sb.Append(p.InnerText);
                            break;
                        case "Heading4":
                            sb.Append("\n#### ");
                            sb.Append(p.InnerText);
                            break;
                        case "Heading5":
                            sb.Append("\n##### ");
                            sb.Append(p.InnerText);
                            break;
                        default:
                            sb.Append("\n\n");
                            sb.Append(p.InnerText);
                            break;
                    }
                }
            }

            return sb.ToString();
        }
    }
}
