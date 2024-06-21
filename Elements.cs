using WorksheetGenerator.Utilities;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using D = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DP = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Packaging;
// using DocumentFormat.OpenXml.Drawing.Charts;

namespace WorksheetGenerator.Elements
{
    public static class El
    {
        public static string? GetImageRelId(Paragraph element)
        {
            Drawing? drawing = element.Descendants<Drawing>().FirstOrDefault();
            D.Blip? blip = drawing?.Descendants<D.Blip>().FirstOrDefault();
            if (blip != null)
            {
                string? embedAttributeValue = blip.Embed?.Value;
                return embedAttributeValue;
            }

            return null;
        }


        // GENERATOR FUNCTIONS

        public static Style Style(
            string id,
            string name,
            string? parentStyle = null,
            ParagraphProperties? pPr = null,
            StyleRunProperties? rPr = null,
            TableProperties? tblPr = null
        )
        {
            Style style = new(
                new AutoRedefine() { Val = OnOffOnlyValues.Off },
                new BasedOn() { Val = "Normal" },
                new LinkedStyle() { Val = "OverdueAmountChar" },
                new Locked() { Val = OnOffOnlyValues.Off },
                new PrimaryStyle() { Val = OnOffOnlyValues.On },
                new StyleHidden() { Val = OnOffOnlyValues.Off },
                new SemiHidden() { Val = OnOffOnlyValues.Off },
                new StyleName() { Val = name },
                new NextParagraphStyle() { Val = "Normal" },
                new UIPriority() { Val = 1 },
                new UnhideWhenUsed() { Val = OnOffOnlyValues.On }
            )
            {
                Type = tblPr != null ? StyleValues.Table : StyleValues.Paragraph,
                StyleId = id,
                CustomStyle = true,
                Default = false
            };

            if (parentStyle != null)
                style.AppendChild(new BasedOn() { Val = parentStyle });

            if (pPr != null)
                style.AppendChild(pPr);
            if (rPr != null)
                style.AppendChild(rPr);
            if (tblPr != null)
                style.AppendChild(tblPr);

            return style;
        }


        public static ParagraphProperties ParagraphStyle(string styleId)
        {
            return new ParagraphProperties(
                new ParagraphStyleId() { Val = styleId }
            );
        }

        public static TableProperties TableStyle(string styleId)
        {
            return new TableProperties(
                new TableStyle() { Val = styleId }
            );
        }

        public static TableBorders TableBorders(EnumValue<BorderValues> val, UInt32Value size, EnumValue<ThemeColorValues> color)
        {
            return new TableBorders(
                new TopBorder()
                {
                    Val = new EnumValue<BorderValues>(val),
                    Size = size,
                    ThemeColor = color
                },
                new BottomBorder()
                {
                    Val = new EnumValue<BorderValues>(val),
                    Size = size,
                    ThemeColor = color
                },
                new LeftBorder()
                {
                    Val = new EnumValue<BorderValues>(val),
                    Size = size,
                    ThemeColor = color
                },
                new RightBorder()
                {
                    Val = new EnumValue<BorderValues>(val),
                    Size = size,
                    ThemeColor = color
                },
                new InsideHorizontalBorder()
                {
                    Val = new EnumValue<BorderValues>(val),
                    Size = size,
                    ThemeColor = color
                },
                new InsideVerticalBorder()
                {
                    Val = new EnumValue<BorderValues>(val),
                    Size = size,
                    ThemeColor = color
                }
            );
        }

        public static TableCellMargin TableCellMargin(int top, int right, int bottom, int left)
        {
            return new TableCellMargin(
                new TopMargin()
                {
                    Width = $"{top}",
                    Type = TableWidthUnitValues.Dxa
                },
                new RightMargin()
                {
                    Width = $"{right}",
                    Type = TableWidthUnitValues.Dxa
                },
                new BottomMargin()
                {
                    Width = $"{bottom}",
                    Type = TableWidthUnitValues.Dxa
                },
                new LeftMargin()
                {
                    Width = $"{left}",
                    Type = TableWidthUnitValues.Dxa
                }
            );
        }

        public static TableCellWidth TableCellWidth(int width)
        {
            return new TableCellWidth()
            {
                Width = $"{width}",
                Type = TableWidthUnitValues.Dxa
            };
        }

        public static Paragraph PageBreak()
        {
            return new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
        }

        public static UInt32Value docPropertiesId = 0;

        public static Paragraph? Image(Paragraph origImage, string relationshipId, Int64Value desiredHeight, string? styleId = null, bool rounded = true)
        {
            docPropertiesId++;

            // Get original image dimensions
            DW.Extent? origExtent = origImage.Descendants<DW.Extent>().FirstOrDefault();
            Int64Value? origCx = origExtent?.Cx;
            Int64Value? origCy = origExtent?.Cy;

            // Get new width
            Int64Value desiredWidth;
            if (origCx != null && origCy != null)
            {
                desiredWidth = HF.GetWidth(origCx, origCy, desiredHeight);

                Drawing drawing =
                    new(
                        new DW.Inline(
                            new DW.Extent() { Cx = desiredWidth, Cy = desiredHeight },
                            new DW.EffectExtent()
                            {
                                LeftEdge = 19050L,
                                TopEdge = 0L,
                                RightEdge = 9525L,
                                BottomEdge = 0L
                            },
                            new DW.DocProperties()
                            {
                                Id = docPropertiesId,
                                Name = "Picture"
                            },
                            new DW.NonVisualGraphicFrameDrawingProperties(
                                new D.GraphicFrameLocks() { NoChangeAspect = true }),
                            new D.Graphic(
                                new D.GraphicData(
                                    new DP.Picture(
                                        new DP.NonVisualPictureProperties(
                                            new DP.NonVisualDrawingProperties()
                                            {
                                                Id = docPropertiesId,
                                                Name = "Image.jpg"
                                            },
                                            new DP.NonVisualPictureDrawingProperties()),
                                        new DP.BlipFill(
                                            new D.Blip(
                                                new D.BlipExtensionList(
                                                    new D.BlipExtension()
                                                    {
                                                        Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                    })
                                            )
                                            {
                                                Embed = relationshipId,
                                                CompressionState = D.BlipCompressionValues.Print
                                            },
                                            new D.Stretch(
                                                new D.FillRectangle())),
                                        new DP.ShapeProperties(
                                            new D.Transform2D(
                                                new D.Offset() { X = 0L, Y = 0L },
                                                new D.Extents() { Cx = desiredWidth, Cy = desiredHeight }),
                                            new D.PresetGeometry(
                                                new D.AdjustValueList()
                                            )
                                            { Preset = rounded ? D.ShapeTypeValues.RoundRectangle : D.ShapeTypeValues.Rectangle }))
                                )
                                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                        )
                        {
                            DistanceFromTop = (UInt32Value)0U,
                            DistanceFromBottom = (UInt32Value)0U,
                            DistanceFromLeft = (UInt32Value)0U,
                            DistanceFromRight = (UInt32Value)0U,
                            EditId = RandomEditId()
                        });

                Paragraph paragraph = new(new Run(drawing));

                if (styleId != null)
                    paragraph.PrependChild(El.ParagraphStyle(styleId));

                return paragraph;
            }

            return null;
        }

        public static string RandomEditId()
        {
            // Generate a random hexadecimal string
            byte[] bytes = new byte[4];
            new Random().NextBytes(bytes);

            // Convert byte array to hexadecimal string
            string randomHex = BitConverter.ToString(bytes).Replace("-", "");

            // Ensure the string has exactly 8 characters
            if (randomHex.Length < 8)
            {
                randomHex = randomHex.PadLeft(8, '0');
            }
            else if (randomHex.Length > 8)
            {
                randomHex = randomHex.Substring(0, 8);
            }

            return randomHex.ToUpper();
        }

        // public static Paragraph NumberListItem(int numId, string text, string styleId = "Text")
        // {
        //     Paragraph item = new Paragraph(
        //         El.ParagraphStyle(styleId),
        //         new Run(new Text(text))
        //     );

        //     ParagraphProperties? pPr = item.Elements<ParagraphProperties>().FirstOrDefault();
        //     pPr?.AppendChild(new NumberingProperties(
        //             new NumberingLevelReference() { Val = 0 },
        //             new NumberingId() { Val = numId }
        //         ));

        //     return item;
        // }

        public static Paragraph ListItem(int numId, string text, string styleId = "Text")
        {
            Paragraph item = new Paragraph(
                El.ParagraphStyle(styleId),
                new Run(new Text(text))
            );

            ParagraphProperties? pPr = item.Elements<ParagraphProperties>().FirstOrDefault();
            pPr?.AppendChild(new NumberingProperties(
                    new NumberingLevelReference() { Val = 0 },
                    new NumberingId() { Val = numId }
                ));

            return item;
        }

        private static int currentListNumId = 1;

        public static List<Paragraph> List(string numberFormatName, MainDocumentPart mainPart, IEnumerable<string> texts, string styleId = "Text", int left = 500, int hanging = 500)
        {
            // Get numbering part
            NumberingDefinitionsPart? numberingPart = mainPart.NumberingDefinitionsPart;
            if (numberingPart == null) throw new InvalidOperationException("NumberingDefinitionsPart is missing.");
            Numbering numbering = numberingPart.Numbering;

            // Number format
            EnumValue<NumberFormatValues> numberFormat;
            if (numberFormatName == "decimal")
                numberFormat = NumberFormatValues.Decimal;
            else if (numberFormatName == "upperLetter")
                numberFormat = NumberFormatValues.UpperLetter;
            else
                numberFormat = NumberFormatValues.Bullet;

            // Create a new abstract numbering definition
            AbstractNum newAbstractNum = new(
                new Level(
                    new NumberingFormat() { Val = numberFormat },
                    new LevelText() { Val = "%1." },
                    new StartNumberingValue() { Val = 1 },
                    new ParagraphProperties(
                        new Indentation() { Left = $"{left}", Hanging = $"{hanging}" }
                    )
                )
                { LevelIndex = 0 }
            )
            { AbstractNumberId = currentListNumId };

            AbstractNum? lastAbstractNum = numbering.Elements<AbstractNum>().LastOrDefault();
            if (lastAbstractNum != null)
                lastAbstractNum.InsertAfterSelf(newAbstractNum);
            else
                numbering.AppendChild(newAbstractNum);

            // Create a new numbering instance
            numbering.AppendChild(new NumberingInstance(
                new AbstractNumId() { Val = currentListNumId }
            )
            { NumberID = currentListNumId }
            );

            // Create list
            List<Paragraph> result = [];
            foreach (string text in texts)
                result.Add(ListItem(currentListNumId, text, styleId));

            currentListNumId++;

            return result;
        }

        public static List<Paragraph> NumberList(MainDocumentPart mainPart, IEnumerable<string> texts, string styleId = "Text", int left = 500, int hanging = 500)
        {
            return List("decimal", mainPart, texts, styleId, left, hanging);
        }

        public static List<Paragraph> LetterList(MainDocumentPart mainPart, IEnumerable<string> texts, string styleId = "Text", int left = 500, int hanging = 500)
        {
            return List("upperLetter", mainPart, texts, styleId, left, hanging);
        }

        public static void KeepNext(Paragraph paragraph)
        {
            ParagraphProperties? pPr = paragraph.Elements<ParagraphProperties>().FirstOrDefault();
            if (pPr == null)
            {
                pPr = new ParagraphProperties();
                paragraph.PrependChild(pPr);
            }

            pPr.Append(new KeepNext());
        }
    }
}