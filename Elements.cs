using System;
using System.Xml.Linq;

namespace WorksheetGenerator.Elements
{
    public static class El
    {
        // NAMESPACES

        public static XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public static XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";


        // HELPER GENERATOR FUNCTIONS

        public static List<XAttribute> TableBorderAttributes(string val, int size, int space, string color, string? themeColor = null)
        {
            List<XAttribute> attributes = [
                new XAttribute(w + "val", val),
                new XAttribute(w + "sz", size),
                new XAttribute(w + "space", space),
                new XAttribute(w + "color", color)
            ];

            if (themeColor != null)
                attributes.Add(new XAttribute(w + "themeColor", themeColor));

            return attributes;
        }


        // GENERATOR FUNCTIONS

        public static XElement Paragraph(string? text = null)
        {
            XElement paragraph = new XElement(w + "p",
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

            SetParagraphFont(paragraph, "Aptos");
            SetParagraphColor(paragraph, "262626", "text1", "D9");

            return paragraph;
        }

        public static XElement Table(
                int num_cols,
                int[] widths,
                List<XAttribute> tableBorderAttributes,
                List<List<XElement>>? paragraphs = null
            )
        {
            widths ??= [11169];
            paragraphs ??= [[]];

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
                        paragraphs[i] == null ? Paragraph() : paragraphs[i]
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
                        new XElement(w + "top", tableBorderAttributes),
                        new XElement(w + "left", tableBorderAttributes),
                        new XElement(w + "bottom", tableBorderAttributes),
                        new XElement(w + "right", tableBorderAttributes),
                        new XElement(w + "insideH", tableBorderAttributes),
                        new XElement(w + "insideV", tableBorderAttributes)
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


        // HELPER STYLING FUNCTIONS

        public static void SetOrAddAttribute(XElement element, XName attrName, object attrVal)
        {
            XAttribute? attribute = element.Attribute(attrName);
            if (attribute == null)
                element.Add(new XAttribute(attrName, attrVal));
            else
                element.SetAttributeValue(attrName, attrVal);
        }

        public static void SetParagraphPropertyStyle(XElement paragraph, XName styleElementName, List<XAttribute> attributes)
        {
            attributes ??= [];

            XElement? pPr = paragraph.Element(w + "pPr");
            XElement? styleElement = pPr?.Element(styleElementName);
            if (styleElement == null)
                pPr?.AddFirst(new XElement(styleElementName, attributes));
            else
                foreach (XAttribute attribute in attributes)
                    SetOrAddAttribute(styleElement, attribute.Name, attribute.Value);
        }

        public static void SetParagraphStyle(XElement paragraph, XName styleElementName, List<XAttribute>? attributes = null)
        {
            attributes ??= [];

            // Edit <w:pPr><w:rPr>
            XElement? pPr = paragraph.Element(w + "pPr");
            XElement? pPr_rPr = pPr?.Element(w + "rPr");
            XElement? styleElement = pPr_rPr?.Element(styleElementName);
            if (styleElement == null)
                pPr_rPr?.AddFirst(new XElement(styleElementName, attributes));
            else
                foreach (XAttribute attribute in attributes)
                    SetOrAddAttribute(styleElement, attribute.Name, attribute.Value);

            // Edit all <w:r><w:rPr>
            IEnumerable<XElement> runs = paragraph.Elements(w + "r");
            foreach (XElement run in runs)
            {
                XElement? r_rPr = run.Element(w + "rPr");
                XElement? r_styleElement = r_rPr?.Element(styleElementName);
                if (r_styleElement == null)
                    r_rPr?.AddFirst(new XElement(styleElementName, attributes));
                else
                    foreach (XAttribute attribute in attributes)
                        SetOrAddAttribute(r_styleElement, attribute.Name, attribute.Value);
            }
        }


        // STYLING FUNCTIONS

        public static void SetParagraphFont(XElement paragraph, string fontName)
        {
            SetParagraphStyle(paragraph, w + "rFonts", [
                new XAttribute(w + "ascii", fontName),
                new XAttribute(w + "hAnsi", fontName)
            ]);
        }

        public static void SetParagraphColor(XElement paragraph, string val, string themeColor, string? themeTint = null)
        {
            List<XAttribute> attributes = [
                new XAttribute(w + "val", val),
                new XAttribute(w + "themeColor", themeColor)
            ];
            if (themeTint != null)
                attributes.Add(new XAttribute(w + "themeTint", themeTint));

            SetParagraphStyle(paragraph, w + "color", attributes);
        }

        public static void SetParagraphSize(XElement paragraph, int size)
        {
            SetParagraphStyle(paragraph, w + "szCs", [
                new XAttribute(w + "val", size)
            ]);

            SetParagraphStyle(paragraph, w + "sz", [
                new XAttribute(w + "val", size)
            ]);
        }

        public static void AddBoldToParagraph(XElement paragraph)
        {
            SetParagraphStyle(paragraph, w + "bCs");
            SetParagraphStyle(paragraph, w + "b");
        }

        public static void CenterParagraph(XElement paragraph)
        {
            SetParagraphPropertyStyle(paragraph, w + "jc", [
                new XAttribute(w + "val", "center")
            ]);
        }

        public static void SetParagraphSpacing(XElement paragraph, int line)
        {
            SetParagraphPropertyStyle(paragraph, w + "spacing", [
                new XAttribute(w + "line", line),
                new XAttribute(w + "lineRule", "auto")
            ]);
        }

        public static void AddTableCellMargin(XElement table, int topSize, int rightSize, int bottomSize, int leftSize)
        {
            XElement? tblPr = table.Element(w + "tblPr");
            XElement? tblCellMar = tblPr?.Element(w + "tblCellMar");
            if (tblCellMar != null)
            {
                XElement? top = tblCellMar.Element(w + "top");
                XElement? right = tblCellMar.Element(w + "right");
                XElement? bottom = tblCellMar.Element(w + "bottom");
                XElement? left = tblCellMar.Element(w + "left");

                if (top != null)
                {
                    SetOrAddAttribute(top, w + "w", topSize);
                    SetOrAddAttribute(top, w + "type", "dxa");
                }
                if (right != null)
                {
                    SetOrAddAttribute(right, w + "w", rightSize);
                    SetOrAddAttribute(right, w + "type", "dxa");
                }
                if (bottom != null)
                {
                    SetOrAddAttribute(bottom, w + "w", bottomSize);
                    SetOrAddAttribute(bottom, w + "type", "dxa");
                }
                if (left != null)
                {
                    SetOrAddAttribute(left, w + "w", leftSize);
                    SetOrAddAttribute(left, w + "type", "dxa");
                }
            }
            else
            {
                tblPr?.Add(new XElement(w + "tblCellMar",
                    new XElement(w + "top",
                        new XAttribute(w + "w", topSize),
                        new XAttribute(w + "type", "dxa")
                    ),
                    new XElement(w + "right",
                        new XAttribute(w + "w", rightSize),
                        new XAttribute(w + "type", "dxa")
                    ),
                    new XElement(w + "bottom",
                        new XAttribute(w + "w", bottomSize),
                        new XAttribute(w + "type", "dxa")
                    ),
                    new XElement(w + "left",
                        new XAttribute(w + "w", leftSize),
                        new XAttribute(w + "type", "dxa")
                    )
                ));
            }
        }
    }
}