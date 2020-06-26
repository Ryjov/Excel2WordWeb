using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Spire.Doc;
using Spire.Xls;
using Spire.Doc.Documents;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel2WordWeb
{
    public class FindAndReplaceObject
    {
        private string wordFilePath;
        private string excelFilePath;
        private string outfilepathfolder;

        public FindAndReplaceObject(string w, string e, string p) { wordFilePath = w; excelFilePath = e; outfilepathfolder = p; }

        public void FindAndReplace()
        {
            string allText = "";
            Document document = new Document();
            document.LoadFromFile($@"{wordFilePath}");
            StringBuilder sb = new StringBuilder();

            foreach (Section section in document.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    sb.AppendLine(paragraph.Text);
                }
            }
            allText = sb.ToString();
            Regex markerRegEx = new Regex(@"<#\d+#[A-Z]+\d+>");
            MatchCollection markerMatches = markerRegEx.Matches(allText);

            Workbook workbook = new Workbook();
            workbook.LoadFromFile($@"C:\Users\Егор\source\repos\Excel2WordWeb\Excel2WordWeb\wwwroot\Files\образец.xls");
            foreach (Match match in markerMatches)
            {
                Regex sheetRegEx = new Regex(@"#\d+#");
                Regex cellRegEx = new Regex(@"#[A-Z]+\d+>");
                int sheetNum = Int32.Parse(sheetRegEx.Match(match.Value).Value.Trim('#'));
                Worksheet sheet = workbook.Worksheets[sheetNum-1];
                string cell = cellRegEx.Match(match.Value).Value.Trim('#', '>');
                string replaceString = sheet.Range[cell].Value;
                document.Replace(match.Value, replaceString, false, false);
            }
            document.SaveToFile($@"{outfilepathfolder}\Files\outfile.docx", Spire.Doc.FileFormat.Docx);
            System.Diagnostics.Process.Start(@"cmd.exe","outfile.docx");


            //System.Diagnostics.Process.Start("ExtractText.txt");
        }
    }
}
