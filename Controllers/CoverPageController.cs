using CoverPage.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using CoverPage.Helpers;
using System.IO;

namespace CoverPage.Controllers
{
    public class CoverPageController : Controller
    {
        public IActionResult Index() => View();

        [HttpPost]
        public IActionResult Download(CoverPageModel model)
        {
            using var stream = new MemoryStream();

            using (var doc = WordprocessingDocument.Create(
                stream,
                WordprocessingDocumentType.Document,
                true))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                ApplyTimesNewRomanStyle(mainPart);
                var body = mainPart.Document.AppendChild(new Body());

                //PAGE SETUP
                body.Append(new SectionProperties(
                    new PageSize
                    {
                        Width = 11906,   // A4
                        Height = 16838
                    },
                    new PageMargin
                    {
                        Top = 1440,     // 1 inch
                        Bottom = 1200,
                        Left = 1440,
                        Right = 1440
                    }
                ));

                // HEADER
                body.Append(TightCenterLine("TRIBHUVAN UNIVERSITY", 26, true));
                body.Append(TightCenterLine("INSTITUTE OF SCIENCE AND TECHNOLOGY", 20, true));
                body.Append(TightCenterLine("AMRIT SCIENCE CAMPUS", 20, true));

                AddImage(
                    mainPart,
                    body,
                    "wwwroot/images/tu-logo.png",
                    width: 200,
                    height: 230
                );
                AddLineImage(mainPart, body);

                body.Append(SmallGap());

                // ================= SUBJECT (Shift+Enter) =================
                body.Append(TightCenterText(
                    new[] { model.SubjectName, "Lab Report" },
                    22,
                    true
                ));

                body.Append(SmallGap());
                // ================= SUBMITTED =================
                body.Append(CreateSubmittedTable(model));

                body.Append(SmallGap());

                // ================= SIGNATURES =================
                body.Append(CreateSignatureTable());
            }

            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "LabReportCover.docx"
            );
        }

        //TEXT HELPERS
        void ApplyTimesNewRomanStyle(MainDocumentPart mainPart)
        {
            var stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylePart.Styles = new Styles();

            // Default "Normal" style
            var normalStyle = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normal",
                Default = true
            };

            normalStyle.Append(
                new StyleName { Val = "Normal" },
                new StyleRunProperties(
                    new RunFonts
                    {
                        Ascii = "Times New Roman",
                        HighAnsi = "Times New Roman",
                        EastAsia = "Times New Roman",
                        ComplexScript = "Times New Roman"
                    },
                    new FontSize { Val = "24" } // 12pt default
                )
            );

            stylePart.Styles.Append(normalStyle);
            stylePart.Styles.Save();
        }
        Paragraph TightCenterLine(string text, int fontSize, bool bold)
        {
            return new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center },
                    new SpacingBetweenLines
                    {
                        Before = "0",
                        After = "0",
                        Line = "220"
                    }
                ),
                new Run(
                    new RunProperties(
                        bold ? new Bold() : null,
                        new FontSize { Val = (fontSize * 2).ToString() }
                    ),
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
                    run.Append(new Break()); // Shift + Enter
            }

            return new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Center },
                    new SpacingBetweenLines
                    {
                        Before = "0",
                        After = "0",
                        Line = "220"
                    }
                ),
                run
            );
        }

        Paragraph InfoLine(string text, bool bold, bool underline = false, bool right = false)
        {
            return new Paragraph(
                new ParagraphProperties(
                    right ? new Justification { Val = JustificationValues.Right } : null,
                    new SpacingBetweenLines
                    {
                        Before = "0",
                        After = "0",
                        Line = "220"
                    }
                ),
                new Run(
                    new RunProperties(
                        bold ? new Bold() : null,
                        underline ? new Underline { Val = UnderlineValues.Single } : null
                    ),
                    new Text(text)
                )
            );
        }

        Paragraph SmallGap()
        {
            return new Paragraph(
                new ParagraphProperties(
                    new SpacingBetweenLines
                    {
                        Before = "0",
                        After = "0",
                        Line = "160"
                    }
                ),
                new Run(new Break())
            );
        }

        // =====================================================
        // ================= IMAGE HELPERS =====================
        // =====================================================

        // 🔥 FIXED: NO LINE SPACING (prevents clipping)
        void AddImage(MainDocumentPart mainPart, Body body, string path, int width, int height)
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using var stream = System.IO.File.OpenRead(path);
            imagePart.FeedData(stream);

            var drawing = ImageHelper.CreateImageDrawing(
                mainPart.GetIdOfPart(imagePart),
                width,
                height
            );

            body.Append(
                new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Center },
                        new SpacingBetweenLines
                        {
                            Before = "200",
                            After = "200"
                        }
                    ),
                    new Run(drawing)
                )
            );
        }

        void AddLineImage(MainDocumentPart mainPart, Body body)
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using var stream = System.IO.File.OpenRead("wwwroot/images/line.png");
            imagePart.FeedData(stream);

            var drawing = ImageHelper.CreateImageDrawing(
                mainPart.GetIdOfPart(imagePart),
                100,
                280
            );

            body.Append(
                new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Center },
                        new SpacingBetweenLines
                        {
                            Before = "160",
                            After = "160"
                        }
                    ),
                    new Run(drawing)
                )
            );
        }

        // =====================================================
        // ================= TABLES ============================
        // =====================================================

        Table CreateSubmittedTable(CoverPageModel model)
        {
            Table table = new Table(
                new TableProperties(
                    new TableWidth { Width = "100%", Type = TableWidthUnitValues.Pct }
                )
            );

            TableRow row = new TableRow();

            row.Append(new TableCell(
                InfoLine("SUBMITTED BY:", true, true),
                InfoLine($"Name: {model.StudentName}", true),
                InfoLine($"Roll: {model.RollNumber}", true),
                InfoLine($"Date: {model.SubmissionDate:yyyy/MM/dd}", true)
            ));

            row.Append(new TableCell(
                InfoLine("SUBMITTED TO:", true, true, true),
                InfoLine(model.TeacherName, true, true, true),
                InfoLine("Department of CSIT", true, false, true)
            ));

            table.Append(row);
            return table;
        }

        Table CreateSignatureTable()
        {
            Table table = new Table(
                new TableProperties(
                    new TableWidth { Width = "100%", Type = TableWidthUnitValues.Pct }
                )
            );

            TableRow row = new TableRow(
                new TableRowProperties(new CantSplit())
            );

            row.Append(new TableCell(
                new Paragraph(new Run(new Text("__________________________"))),
                new Paragraph(new Run(new Text("External Teacher's Signature")))
            ));

            row.Append(new TableCell(
                new Paragraph(
                    new ParagraphProperties(new Justification { Val = JustificationValues.Right }),
                    new Run(new Text("__________________________"))
                ),
                new Paragraph(
                    new ParagraphProperties(new Justification { Val = JustificationValues.Right }),
                    new Run(new Text("Internal Teacher's Signature"))
                )
            ));

            table.Append(row);
            return table;
        }
    }
}
