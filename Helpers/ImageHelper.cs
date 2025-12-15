using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace CoverPage.Helpers
{
    public static class ImageHelper
    {
        public static Drawing CreateImageDrawing(string relationshipId, long widthPx, long heightPx)
        {
            long widthEmu = widthPx * 9525;
            long heightEmu = heightPx * 9525;

            return new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                    new DW.EffectExtent
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.DocProperties
                    {
                        Id = 1U,
                        Name = "TU Logo"
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new GraphicFrameLocks { NoChangeAspect = true }
                    ),
                    new Graphic(
                        new GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties
                                    {
                                        Id = 0U,
                                        Name = "logo.png"
                                    },
                                    new PIC.NonVisualPictureDrawingProperties()
                                ),
                                new PIC.BlipFill(
                                    new Blip
                                    {
                                        Embed = relationshipId,
                                        CompressionState = BlipCompressionValues.Print
                                    },
                                    new Stretch(new FillRectangle())
                                ),
                                new PIC.ShapeProperties(
                                    new Transform2D(
                                        new Offset { X = 0, Y = 0 },
                                        new Extents { Cx = widthEmu, Cy = heightEmu }
                                    ),
                                    new PresetGeometry(
                                        new AdjustValueList()
                                    )
                                    { Preset = ShapeTypeValues.Rectangle }
                                )
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
            );
        }
    }
}
