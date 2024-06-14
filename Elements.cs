using System;
using System.Xml.Linq;

namespace WorksheetGenerator.Elements
{
    public static class El
    {
        public static XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public static XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

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

        public static void SetOrAddAttribute(XElement element, XName attrName, object attrVal)
        {
            XAttribute? attribute = element.Attribute(attrName);
            if (attribute == null)
                element.Add(new XAttribute(attrName, attrVal));
            else
                element.SetAttributeValue(attrName, attrVal);
        }

        public static void SetParagraphSize(XElement paragraph, int size)
        {
            // Edit <w:pPr><w:rPr>
            XElement? pPr = paragraph.Element(w + "pPr");
            XElement? pPr_rPr = pPr?.Element(w + "rPr");
            XElement? szCs = pPr_rPr?.Element(w + "szCs");
            if (szCs == null)
                pPr_rPr?.AddFirst(new XElement(w + "szCs",
                    new XAttribute(w + "val", size)
                ));
            else
                SetOrAddAttribute(szCs, w + "val", size);

            XElement? sz = pPr_rPr?.Element(w + "sz");
            if (sz == null)
                pPr_rPr?.AddFirst(new XElement(w + "sz",
                    new XAttribute(w + "val", size)
                ));
            else
                SetOrAddAttribute(sz, w + "val", size);

            // Edit all <w:r><w:rPr>
            IEnumerable<XElement> runs = paragraph.Elements(w + "r");
            foreach (XElement run in runs)
            {
                XElement? r_rPr = run.Element(w + "rPr");
                XElement? r_rPr_szCs = pPr_rPr?.Element(w + "szCs");
                if (r_rPr_szCs == null)
                    r_rPr?.AddFirst(new XElement(w + "szCs",
                        new XAttribute(w + "val", size)
                    ));
                else
                    SetOrAddAttribute(r_rPr_szCs, w + "val", size);

                XElement? r_rPr_sz = r_rPr?.Element(w + "sz");
                if (r_rPr_sz == null)
                    r_rPr?.AddFirst(new XElement(w + "sz",
                        new XAttribute(w + "val", size)
                    ));
                else
                    SetOrAddAttribute(r_rPr_sz, w + "val", size);
            }
        }

        public static void AddBoldToRunProperty(XElement runProperty)
        {
            XElement? bCs = runProperty?.Element(w + "bCs");
            if (bCs == null)
                runProperty?.AddFirst(new XElement(w + "bCs"));
            XElement? b = runProperty?.Element(w + "b");
            if (b == null)
                runProperty?.AddFirst(new XElement(w + "b"));
        }

        public static void AddBoldToParagraph(XElement paragraph)
        {
            XElement? pPr = paragraph.Element(w + "pPr");
            XElement? pPr_rPr = pPr?.Element(w + "rPr");
            if (pPr_rPr != null)
                AddBoldToRunProperty(pPr_rPr);

            IEnumerable<XElement> runs = paragraph.Elements(w + "r");
            foreach (XElement run in runs)
            {
                XElement? r_rPr = run.Element(w + "rPr");
                if (r_rPr != null)
                    AddBoldToRunProperty(r_rPr);
            }
        }

        public static void CenterParagraph(XElement paragraph)
        {
            XElement? pPr = paragraph.Element(w + "pPr");

            XElement? jc = pPr?.Element(w + "jc");
            if (jc == null)
                pPr?.Add(new XElement(w + "jc"));
            jc = pPr?.Element(w + "jc");

            XAttribute? val = jc?.Attribute(w + "val");
            if (val == null)
                jc?.Add(new XAttribute(w + "val", "center"));
            else
                jc?.SetAttributeValue(w + "val", "center");
        }

        public static void AddTitleStyles(XElement paragraph)
        {
            CenterParagraph(paragraph);
            SetParagraphSize(paragraph, 36);
            AddBoldToParagraph(paragraph);
            Console.WriteLine(paragraph);
        }

        public static List<XAttribute> TableBorderAttributes(string val, int size, int space, string color)
        {
            return [
                new XAttribute(w + "val", val),
                new XAttribute(w + "sz", size),
                new XAttribute(w + "space", space),
                new XAttribute(w + "color", color)
            ];
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
                        new XElement(w + "top", TableBorderAttributes("none", 0, 0, "auto")),
                        new XElement(w + "left", TableBorderAttributes("none", 0, 0, "auto")),
                        new XElement(w + "bottom", TableBorderAttributes("none", 0, 0, "auto")),
                        new XElement(w + "right", TableBorderAttributes("none", 0, 0, "auto")),
                        new XElement(w + "insideH", TableBorderAttributes("none", 0, 0, "auto")),
                        new XElement(w + "insideV", TableBorderAttributes("none", 0, 0, "auto"))
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