using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordManipulation
{
    class Program
    {
        static void Main(string[] args)
        {
            string destination = "c:\\users\\mohd\\documents\\visual studio 2013\\Projects\\WordManipulation\\WordManipulation\\AppData\\Sample.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.Open(destination, true))
            {
                var mainDocPart = doc.MainDocumentPart;
                if (doc == null)
                {
                    mainDocPart = doc.AddMainDocumentPart();
                }

                if (mainDocPart.Document == null)
                {
                    mainDocPart.Document = new Document();
                }

                ApplyHeader(doc);

                ApplyFooter(doc);
                
            }
        }
        public static void ApplyHeader(WordprocessingDocument doc)
        {
            // Get the main document part.
            MainDocumentPart mainDocPart = doc.MainDocumentPart;

            HeaderPart headerPart1 = mainDocPart.AddNewPart<HeaderPart>("r97");
            Header header1 = new Header();
            Paragraph paragraph1 = new Paragraph() { };
            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = "Header stuff";
            run1.Append(text1);
            paragraph1.Append(run1);
            header1.Append(paragraph1);
            headerPart1.Header = header1;
            SectionProperties sectionProperties1 = mainDocPart.Document.Body.Descendants<SectionProperties>().FirstOrDefault();
            if (sectionProperties1 == null)
            {
                sectionProperties1 = new SectionProperties() { };
                mainDocPart.Document.Body.Append(sectionProperties1);
            }
            HeaderReference headerReference1 = new HeaderReference() { Type = HeaderFooterValues.Default, Id = "r97" };
            sectionProperties1.InsertAt(headerReference1, 0);

        }

        public static void ApplyFooter(WordprocessingDocument doc)
        {
            // Get the main document part.
            MainDocumentPart mainDocPart = doc.MainDocumentPart;
            FooterPart footerPart1 = mainDocPart.AddNewPart<FooterPart>("r98");
            Footer footer1 = new Footer();
            Paragraph paragraph1 = new Paragraph() { };
            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = "Imran 172.100.10.10 1st may 2016";

            run1.Append(text1);
            paragraph1.Append(run1);
            footer1.Append(paragraph1);
            footerPart1.Footer = footer1;
            SectionProperties sectionProperties1 = mainDocPart.Document.Body.Descendants<SectionProperties>().FirstOrDefault();
            if (sectionProperties1 == null)
            {
                sectionProperties1 = new SectionProperties() { };
                mainDocPart.Document.Body.Append(sectionProperties1);
            }
            FooterReference footerReference1 = new FooterReference() { Type = DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default, Id = "r98" };
            sectionProperties1.InsertAt(footerReference1, 0);

        }
    }
}
