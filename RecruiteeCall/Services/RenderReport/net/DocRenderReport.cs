using RecruiteeCall.Models;
using System.Collections.Generic;
using System.IO;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace RecruiteeCall.Services.RenderReport.net
{
    public class DocRenderReport : IRenderReport
    {
        public string getReport(int candidateId, string candidateName, string templateName, IList<Question> questions)
        {
            string fileName = $"{Path.GetTempPath()}\\Candidate {candidateId}.docx";
            using (DocX document = DocX.Create(fileName))
            {
                document.SetDefaultFont(new Font("Arial"));

                // Generate the Headers/Footers for this document
                document.AddHeaders();
                document.AddFooters();
                // Insert a Paragraph in the Headers/Footers
                string headerImagePath = "./Images/Header.jpg";
                Image headerImage = document.AddImage(headerImagePath);
                Picture headerPicture = headerImage.CreatePicture();

                string FooterImagePath = "./Images/Footer.jpg";
                Image footerImage = document.AddImage(FooterImagePath);
                Picture footerPicture = footerImage.CreatePicture();

                document.Headers.Even.InsertParagraph(candidateName);
                document.Headers.Even.InsertParagraph(templateName);
                var p = document.Headers.Even.InsertParagraph();
                p.AppendPicture(headerPicture);
                p.Alignment = Alignment.right;
                document.Headers.Odd.InsertParagraph(candidateName);
                document.Headers.Odd.InsertParagraph(templateName);
                p = document.Headers.Odd.InsertParagraph();
                p.AppendPicture(headerPicture);
                p.Alignment = Alignment.right;

                p = document.Footers.Even.InsertParagraph();
                footerPicture.Width = footerPicture.Width * .7f;
                footerPicture.Height = footerPicture.Height * .7f;
                p.AppendPicture(footerPicture);
                document.Footers.Even.InsertParagraph("1 university ave, 3rd floor, toronto, ontario  canada m5j 2p1,  416.599.0000  paralucent.com");
                p = document.Footers.Odd.InsertParagraph();
                p.AppendPicture(footerPicture);
                p.Alignment = Alignment.center;
                document.Footers.Odd.InsertParagraph("1 university ave, 3rd floor, toronto, ontario  canada m5j 2p1,  416.599.0000  paralucent.com");

                foreach (var q in questions)
                {
                    document.InsertParagraph($"{q.Index}. {q.Title}").Bold();
                    document.InsertParagraph(q.Content);
                    document.InsertParagraph();
                }

                // Save the document.
                document.Save();
            }
            return fileName;
        }
    }
}
