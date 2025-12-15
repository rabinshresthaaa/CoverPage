using CoverPage.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using CoverPage.Helpers;
using Microsoft.AspNetCore.Razor.TagHelpers;
using System.IO;

namespace CoverPage.Controllers
{
    public class CoverPageController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Download(CoverPageModel model)
        {
            return GenerateDoc(model);
        }

        private IActionResult GenerateDoc(CoverPageModel model)
        {
            using var stream = new MemoryStream();

            using (var doc = WordprocessingDocument.Create(
                stream,
                WordprocessingDocumentType.Document,
                true))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                // A4 Page Setup
                body.Append(new SectionProperties(
                    new PageSize { Width = 11906, Height = 16838 },
                    new PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 }
                ));

                // UNIVERSITY HEADER
                body.Append(CenterText("TRIBHUVAN UNIVERSITY", 28, true));
                body.Append(CenterText("INSTITUTE OF SCIENCE AND TECHNOLOGY", 24, true));
                body.Append(CenterText("AMRIT SCIENCE CAMPUS", 22, true));

                body.Append(EmptyLine(2));

                // LOGO
                AddImage(mainPart, body);

                body.Append(EmptyLine(2));

                // SUBJECT (USER INPUT)
                body.Append(CenterText(model.SubjectName, 22, true));
                body.Append(CenterText("Lab Report", 20, false));

                body.Append(EmptyLine(3));

                // SUBMITTED BY / TO
                body.Append(CreateSubmittedTable(model));

                body.Append(EmptyLine(3));

                // SIGNATURES
                body.Append(CreateSignatureTable());
            }

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "LabReportCover.docx");
        }

        Paragraph CenterText(string text, int size, bool bold)
        {
            return new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center }),
                new Run(
                    new RunProperties(
                        new Bold { Val = bold },
                        new FontSize { Val = (size * 2).ToString() }),
                    new Text(text)
                )
            );
        }

        Paragraph EmptyLine(int count)
        {
            var p = new Paragraph();
            for (int i = 0; i < count; i++)
                p.Append(new Run(new Break()));
            return p;
        }

        Table CreateSubmittedTable(CoverPageModel model)
        {
            Table table = new Table();
            TableRow row = new TableRow();

            // LEFT COLUMN
            row.Append(new TableCell(
                new Paragraph(
                    new Run(
                        new RunProperties(new Bold()),
                        new Text("SUBMITTED BY:")
                    )
                ),
                new Paragraph(new Run(new Text($"Name: {model.StudentName}"))),
                new Paragraph(new Run(new Text($"Roll: {model.RollNumber}"))),
                new Paragraph(new Run(new Text($"Date: {model.SubmissionDate:yyyy/MM/dd}")))
            ));

            // RIGHT COLUMN
            row.Append(new TableCell(
                new Paragraph(
                    new Run(
                        new RunProperties(new Bold()),
                        new Text("SUBMITTED TO:")
                    )
                ),
                new Paragraph(new Run(new Text(model.TeacherName))),
                new Paragraph(new Run(new Text("Department of CSIT")))
            ));

            table.Append(row);
            return table;
        }

        Table CreateSignatureTable()
        {
            Table table = new Table();
            TableRow row = new TableRow();

            row.Append(new TableCell(
                new Paragraph(new Run(new Text("__________________________"))),
                new Paragraph(new Run(new Text("External Teacher's Signature")))
            ));

            row.Append(new TableCell(
                new Paragraph(new Run(new Text("__________________________"))),
                new Paragraph(new Run(new Text("Internal Teacher's Signature")))
            ));

            table.Append(row);
            return table;
        }

        void AddImage(MainDocumentPart mainPart, Body body)
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);

            using var stream = System.IO.File.OpenRead("wwwroot/images/tu-logo.png");
            imagePart.FeedData(stream);

            var drawing = ImageHelper.CreateImageDrawing(
                mainPart.GetIdOfPart(imagePart), 200, 200);

            body.Append(new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center }),
                new Run(drawing)
            ));
        }


    }
}
