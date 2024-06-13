using System;
using System.Xml.Linq;

namespace WorksheetGenerator.Elements
{
    public static class El
    {
        public static XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static XElement titleRunProperty =
            new XElement(w + "rPr",
                new XElement(w + "b"),
                new XElement(w + "bCs"),
                new XElement(w + "sz",
                    new XAttribute(w + "val", "36")
                ),
                new XElement(w + "szCs",
                    new XAttribute(w + "val", "36")
                ),
                new XElement(w + "lang",
                    new XAttribute(w + "val", "en-US")
                )
            );

        public static XElement ParagraphElement(string? text = null)
        {
            return new XElement(w + "p",
                new XElement(w + "pPr",
                    new XElement(w + "rPr",
                        new XElement(w + "lang",
                            new XAttribute(w + "val", "en-US")
                        )
                    )
                ),
                text == null ? null : new XElement(w + "r",
                    new XElement(w + "rPr",
                        new XElement(w + "lang",
                            new XAttribute(w + "val", "en-US")
                        )
                    ),
                    new XElement(w + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        text
                    )
                )
            );
        }

        public static XElement TableElement(int num_cols = 1, int[]? widths = null, List<List<XElement>>? paragraphs = null)
        {
            if (widths == null)
                widths = [9350];

            if (paragraphs == null)
                paragraphs = [[]];

            // Create columns
            List<XElement> columnElements = [];
            for (int i = 0; i < num_cols; i++)
            {
                columnElements.Add(
                    new XElement(w + "tc",
                        new XElement(w + "tcPr",
                            new XElement(w + "tcW",
                                new XAttribute(w + "w", widths[i]),
                                new XAttribute(w + "type", "dxa")
                            )
                        ),
                        // Ignore ID attributes for now
                        paragraphs[i] != null ? paragraphs[i] : ParagraphElement()
                    )
                );
            }

            // Create table
            return new XElement(w + "tbl",
                new XElement(w + "tblPr",
                    new XElement(w + "tblStyle",
                        new XAttribute(w + "val", "TableGrid")
                    ),
                    new XElement(w + "tblW",
                        new XAttribute(w + "w", "0"),
                        new XAttribute(w + "type", "auto")
                    ),
                    new XElement(w + "tblBorders",
                        new XElement(w + "top",
                            new XAttribute(w + "val", "none"),
                            new XAttribute(w + "sz", "0"),
                            new XAttribute(w + "space", "0"),
                            new XAttribute(w + "color", "auto")
                        ),
                        new XElement(w + "left",
                            new XAttribute(w + "val", "none"),
                            new XAttribute(w + "sz", "0"),
                            new XAttribute(w + "space", "0"),
                            new XAttribute(w + "color", "auto")
                        ),
                        new XElement(w + "bottom",
                            new XAttribute(w + "val", "none"),
                            new XAttribute(w + "sz", "0"),
                            new XAttribute(w + "space", "0"),
                            new XAttribute(w + "color", "auto")
                        ),
                        new XElement(w + "right",
                            new XAttribute(w + "val", "none"),
                            new XAttribute(w + "sz", "0"),
                            new XAttribute(w + "space", "0"),
                            new XAttribute(w + "color", "auto")
                        ),
                        new XElement(w + "insideH",
                            new XAttribute(w + "val", "none"),
                            new XAttribute(w + "sz", "0"),
                            new XAttribute(w + "space", "0"),
                            new XAttribute(w + "color", "auto")
                        ),
                        new XElement(w + "insideV",
                            new XAttribute(w + "val", "none"),
                            new XAttribute(w + "sz", "0"),
                            new XAttribute(w + "space", "0"),
                            new XAttribute(w + "color", "auto")
                        )
                    ),
                    new XElement(w + "tblLook",
                        new XAttribute(w + "val", "04A0"),
                        new XAttribute(w + "firstRow", "1"),
                        new XAttribute(w + "lastRow", "0"),
                        new XAttribute(w + "firstColumn", "1"),
                        new XAttribute(w + "lastColumn", "0"),
                        new XAttribute(w + "noHBand", "0"),
                        new XAttribute(w + "noVBand", "1")
                    )
                ),
                new XElement(w + "tblGrid",
                    new XElement(w + "gridCol",
                        new XAttribute(w + "w", "6374")
                    ),
                    new XElement(w + "gridCol",
                        new XAttribute(w + "w", "2976")
                    )
                ),
                new XElement(w + "tr", // Ignoring ID attributes for now
                    columnElements
                )
            );
        }
    }
}