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

                // UNIVERSITY HEADERS
                body.Append(TightCenterText(new[]
                {
                    "TRIBHUVAN UNIVERSITY",
                    "INSTITUTE OF SCIENCE AND TECHNOLOGY",
                    "AMRIT SCIENCE CAMPUS"
                }, 26, true));


                body.Append(EmptyLine(1));

                // LOGO
                AddImage(mainPart, body);

                body.Append(EmptyLine(1));

                // SUBJECT (USER INPUT)
                body.Append(CenterText(model.SubjectName, 22, true));
                body.Append(CenterText("Lab Report", 20, false));

                body.Append(EmptyLine(1));

                // SUBMITTED BY / TO
                body.Append(CreateSubmittedTable(model));

                body.Append(EmptyLine(1));

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
        Paragraph TightCenterText(string[] lines, int fontSize, bool bold)
        {
            var run = new Run(
                new RunProperties(
                    bold ? new Bold() : null,
                    new FontSize { Val = (fontSize * 2).ToString() }
                )
            );

            for (int i = 0; i < lines.Length; i++)
            {
                run.Append(new Text(lines[i]));
                if (i < lines.Length - 1)
                    run.Append(new Break()); // Shift+Enter
            }

            return new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center }
                ),
                run
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

            // Make table full width
            table.AppendChild(new TableProperties(
                new TableWidth { Width = "100%", Type = TableWidthUnitValues.Pct }
            ));

            TableRow row = new TableRow();

            // LEFT CELL
            row.Append(new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Width = "50%", Type = TableWidthUnitValues.Pct }
                ),
                new Paragraph(
                    new Run(new RunProperties(new Bold()), new Text("SUBMITTED BY:"))
                ),
                new Paragraph(new Run(new Text($"Name: {model.StudentName}"))),
                new Paragraph(new Run(new Text($"Roll: {model.RollNumber}"))),
                new Paragraph(new Run(new Text($"Date: {model.SubmissionDate:yyyy/MM/dd}")))
            ));

            // RIGHT CELL
            row.Append(new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Width = "50%", Type = TableWidthUnitValues.Pct }
                ),
                new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Right }
                    ),
                    new Run(new RunProperties(new Bold()), new Text("SUBMITTED TO:"))
                ),
                new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Right }
                    ),
                    new Run(new Text(model.TeacherName))
                ),
                new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Right }
                    ),
                    new Run(new Text("Department of CSIT"))
                )
            ));

            table.Append(row);
            return table;
        }


        Table CreateSignatureTable()
        {
            Table table = new Table();

            table.AppendChild(new TableProperties(
                new TableWidth { Width = "100%", Type = TableWidthUnitValues.Pct },
                new TableLook { Val = "04A0" }
            ));

            TableRow row = new TableRow(
                new TableRowProperties(
                    new CantSplit() // 🔥 prevents page break
                )
            );

            row.Append(new TableCell(
                new Paragraph(new Run(new Text("__________________________"))),
                new Paragraph(new Run(new Text("External Teacher's Signature")))
            ));

            row.Append(new TableCell(
                new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Right }
                    ),
                    new Run(new Text("__________________________"))
                ),
                new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Right }
                    ),
                    new Run(new Text("Internal Teacher's Signature"))
                )
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
