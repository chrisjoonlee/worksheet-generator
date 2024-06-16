using System;
using System.Xml.Linq;
using WorksheetGenerator.Elements;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WorksheetGenerator.Utilities
{
    public static class HF
    {
        // public static XElement? GetElementOnly(XElement? element)
        // {
        //     if (element != null)
        //     {
        //         // Create a new element with the same name and attributes
        //         XElement newElement = new XElement(element.Name);
        //         foreach (XAttribute attribute in element.Attributes())
        //             newElement.Add(new XAttribute(attribute));
        //         return newElement;
        //     }
        //     else
        //         return null;
        // }

        // public static XElement? GetDocumentAndBodyOnly(XDocument doc)
        // {
        //     XElement? originalDocElement = doc.Element(El.w + "document");
        //     XElement? docElement = GetElementOnly(doc.Element(El.w + "document"));
        //     if (docElement != null && originalDocElement != null)
        //     {
        //         XElement? body = GetElementOnly(originalDocElement.Element(El.w + "body"));
        //         if (body != null)
        //         {
        //             docElement.Add(body);
        //             return docElement;
        //         }
        //         else
        //             return null;
        //     }
        //     else
        //         return null;
        // }

        public static bool IsIdentifier(OpenXmlElement element)
        {
            string text = GetParagraphText(element);

            if (string.IsNullOrEmpty(text))
                return false;

            foreach (char c in text)
                if (char.IsLetter(c) && !char.IsUpper(c))
                    return false;

            return true;
        }

        public static string GetParagraphText(OpenXmlElement paragraph)
        {
            if (paragraph is Paragraph)
                return string.Concat(paragraph.Descendants<Text>().Select(t => t.Text)).Trim();
            else
                return "";
        }

        public static bool ElementTextStartsWith(OpenXmlElement element, string str)
        {
            // Console.WriteLine(GetParagraphText(element));
            return GetParagraphText(element).StartsWith(str, StringComparison.CurrentCultureIgnoreCase);
        }

        public static string RemovePrefix(string text)
        {
            // Find the index of the colon
            int colonIndex = text.IndexOf(':');

            // If colon is found and it's not the last character
            if (colonIndex != -1 && colonIndex < text.Length - 1)
            {
                // Extract and return the substring after the colon (excluding colon and leading whitespace)
                return text.Substring(colonIndex + 1).TrimStart();
            }

            // Return the original string if colon is not found or it's the last character
            return text;
        }

        public static Paragraph? GetWorksheetTitleElement(OpenXmlElementList origElementList)
        {
            foreach (OpenXmlElement element in origElementList)
            {
                if (ElementTextStartsWith(element, "title:"))
                {
                    string title = RemovePrefix(GetParagraphText(element)).ToUpper();
                    Paragraph worksheetTitlePara = new(
                        new ParagraphProperties(
                            new ParagraphStyleId() { Val = "WorksheetTitle" }
                            // new Justification() { Val = JustificationValues.Center }
                        ),
                        new Run(
                            new RunProperties(
                            // new FontSize() { Val = "48" },
                            // new FontSizeComplexScript() { Val = "48" },
                            // new Bold(),
                            // new BoldComplexScript(),
                            // new RunFonts() { Ascii = "Aptos" },
                            // new Color() { Val = "262626" }
                            ),
                            new Text(title)
                        )
                    );

                    // AddWorksheetTitleStyles(worksheetTitlePara);
                    return worksheetTitlePara;
                }
            }
            return null;
        }

        // public static bool ContainsIdentifier(IEnumerable<XElement> paragraphs, string identifierName)
        // {
        //     foreach (XElement paragraph in paragraphs)
        //         if (IsIdentifier(paragraph) && ((string)paragraph).StartsWith(identifierName))
        //             return true;
        //     return false;
        // }

        public static bool HasText(OpenXmlElement element)
        {
            return !string.IsNullOrWhiteSpace(GetParagraphText(element));
        }

        public static List<OpenXmlElement> GetParagraphsByIdentifier(OpenXmlElementList elements, string identifierName)
        {
            bool isBetweenIdentifiers = false;
            List<OpenXmlElement> result = [];

            foreach (OpenXmlElement element in elements)
            {
                if (IsIdentifier(element))
                {
                    if (isBetweenIdentifiers)
                        break;

                    if (ElementTextStartsWith(element, identifierName))
                    {
                        isBetweenIdentifiers = true;
                        result = [];
                    }
                }
                else if (isBetweenIdentifiers)
                    if (!ElementTextStartsWith(element, "chatgpt:") && HasText(element))
                        result.Add(element);
            }

            return result;
        }

        // public static bool IsImage(XElement element)
        // {
        //     foreach (XElement child in element.Elements())
        //     {
        //         // Check if the child element is a drawing or picture
        //         if (child.Name == El.w + "drawing" || child.Name == El.w + "pict")
        //             return true;

        //         // Check if the child element is a run element containing drawing or picture
        //         if (child.Name == El.w + "r")
        //             foreach (XElement runChild in child.Elements())
        //                 if (runChild.Name == El.w + "drawing" || runChild.Name == El.w + "pict")
        //                     return true;
        //     }

        //     return false;
        // }

        // public static ulong GetWidth(double width, double height, double desiredHeight)
        // {
        //     return (ulong)Math.Round((double)(desiredHeight / height * width));
        // }

        // public static void FormatImage(XElement element, double desiredHeight = 1640000)
        // {
        //     // Add elements to <w:pPr>
        //     XElement? paragraphProperty = element.Element(El.w + "pPr");
        //     paragraphProperty?.AddFirst(new XElement(El.w + "jc",
        //         new XAttribute(El.w + "val", "center")
        //     ));
        //     paragraphProperty?.AddFirst(new XElement(El.w + "spacing",
        //         new XAttribute(El.w + "before", 240),
        //         new XAttribute(El.w + "line", 400),
        //         new XAttribute(El.w + "lineRule", "auto")
        //     ));

        //     // Add elements to <w:rPr>
        //     XElement? runProperty = paragraphProperty?.Element(El.w + "rPr");
        //     runProperty?.AddFirst(new XElement(El.w + "szCs",
        //         new XAttribute(El.w + "val", 36)
        //     ));
        //     runProperty?.AddFirst(new XElement(El.w + "sz",
        //         new XAttribute(El.w + "val", 36)
        //     ));
        //     runProperty?.AddFirst(new XElement(El.w + "bCs"));
        //     runProperty?.AddFirst(new XElement(El.w + "b"));

        //     // Edit <wp:extent> (Resizing)
        //     XElement? extentElement = element.Descendants(El.wp + "extent").FirstOrDefault();
        //     XAttribute? cx = extentElement?.Attribute("cx");
        //     XAttribute? cy = extentElement?.Attribute("cy");
        //     if (cx != null && cy != null)
        //     {
        //         bool validCx = double.TryParse(cx.Value, out double origWidth);
        //         bool validCy = double.TryParse(cy.Value, out double origHeight);

        //         if (validCx && validCy)
        //         {
        //             double desiredWidth = GetWidth(origWidth, origHeight, desiredHeight);
        //             extentElement?.SetAttributeValue("cx", desiredWidth);
        //             extentElement?.SetAttributeValue("cy", desiredHeight);

        //             // Edit <a:xfrm><a:ext> (Resizing)
        //             XElement? transformElement = element.Descendants(El.a + "xfrm").FirstOrDefault();
        //             XElement? extElement = transformElement?.Element(El.a + "ext");

        //             extElement?.SetAttributeValue("cx", desiredWidth);
        //             extElement?.SetAttributeValue("cy", desiredHeight);
        //         }
        //     }

        //     // Edit <a:prstGeom> (Rounded corners)
        //     XElement? geometryElement = element.Descendants(El.a + "prstGeom").FirstOrDefault();
        //     geometryElement?.SetAttributeValue("prst", "roundRect");
        // }

        // public static void AddSectionTitleStyles(XElement paragraph)
        // {
        //     El.CenterParagraph(paragraph);
        //     El.SetParagraphSize(paragraph, 36);
        //     El.AddBoldToParagraph(paragraph);
        //     El.SetParagraphColor(paragraph, "0F9ED5", "accent4");
        // }

        // public static void AddWorksheetTitleStyles(Paragraph paragraph)
        // {
        //     El.CenterParagraph(paragraph);
        //     El.SetParagraphSize(paragraph, 48);
        //     El.AddBoldToParagraph(paragraph);
        // }

        // public static XElement? GetSectionTitleElement(IEnumerable<XElement> elements)
        // {
        //     foreach (XElement element in elements)
        //         if (StartsWith(element, "title:"))
        //             return element;

        //     return null;
        // }

        public static Paragraph GetFormattedSectionTitleElement(string title, int sectionNo = -1)
        {
            string formattedTitle;
            if (sectionNo >= 0)
                formattedTitle = sectionNo + ". " + title.Trim();
            else
                formattedTitle = title.Trim();

            Paragraph titleElement = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId() { Val = "SectionTitle" }
                ),
                new Run(
                    new Text(formattedTitle)
                )
            );
            return titleElement;
        }

        // public static XElement GetFormattedAnswerKeySectionTitleElement(string title, int sectionNo = -1)
        // {
        //     string formattedTitle;
        //     if (sectionNo >= 0)
        //         formattedTitle = sectionNo + ". " + title.Trim().ToUpper();
        //     else
        //         formattedTitle = title.Trim().ToUpper();

        //     XElement titleElement = El.Paragraph(formattedTitle);
        //     El.AddBoldToParagraph(titleElement);
        //     return titleElement;
        // }

        public static Dictionary<string, string> GetVocab(List<OpenXmlElement> elements)
        {
            Dictionary<string, string> vocab = [];

            foreach (OpenXmlElement element in elements)
            {
                string text = GetParagraphText(element);

                int colonIndex = text.IndexOf(':');
                if (colonIndex != -1 && colonIndex < text.Length - 1)
                {
                    string word = text.Substring(0, colonIndex);
                    string definition = text.Substring(colonIndex + 1).TrimStart();
                    // Remove any final periods
                    if (definition[definition.Length - 1] == '.')
                        definition = definition.Substring(0, definition.Length - 1);
                    vocab.Add(word, definition);
                }
            }

            return vocab;
        }

        public static OpenXmlElement VocabBox(ICollection<string> words)
        {
            string formattedWords = string.Join("        ", words);
            Paragraph formattedWordsPara = new(
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center }
                ),
                new Run(
                    new RunProperties(
                    new FontSize() { Val = "28" },
                    new FontSizeComplexScript() { Val = "28" },
                    new Bold()
                ),
                    new Text(formattedWords)
                )
            );
            // El.SetParagraphLine(paragraph, 440);
            // El.SetParagraphSpacing(paragraph, 0, 200);
            List<List<OpenXmlElement>> content = [[formattedWordsPara]];

            int[] tableColumnWidths = [11169];
            // XElement box = El.Table(
            //     tableColumnWidths,
            //     El.TableBorderAttributes("single", 24, 0, "0F9ED5", "accent4")
            // );
            // box.Add(El.TableRow(tableColumnWidths, content));

            // El.AddTableCellMargin(box, 440, 440, 0, 440);

            // return box;
            return formattedWordsPara;
        }

        // public static Dictionary<TKey, TValue> ShuffledDictionary<TKey, TValue>(Dictionary<TKey, TValue> dict) where TKey : notnull
        // {
        //     Random random = new();

        //     // Convert the dictionary to a list of key-value pairs
        //     List<KeyValuePair<TKey, TValue>> keyValuePairs = dict.ToList();

        //     // Shuffle the list using the Fisher-Yates algorithm
        //     int n = keyValuePairs.Count;
        //     while (n > 1)
        //     {
        //         n--;
        //         int k = random.Next(n + 1);
        //         (keyValuePairs[n], keyValuePairs[k]) = (keyValuePairs[k], keyValuePairs[n]);
        //     }

        //     // Create a new dictionary from the shuffled list
        //     return keyValuePairs.ToDictionary(pair => pair.Key, pair => pair.Value);
        // }

        // public static (XElement, List<XElement>) VocabBlanksAndDefinitions(Dictionary<string, string> vocab)
        // {
        //     int[] tableColWidths = [3261, 7938];

        //     // Create table
        //     XElement mainTable = El.Table(
        //         tableColWidths,
        //         El.TableBorderAttributes("none", 0, 0, "auto")
        //     );

        //     // Create answer list
        //     List<XElement> answerList = [];

        //     // Shuffle words
        //     Dictionary<string, string> shuffledVocab = ShuffledDictionary(vocab);

        //     // Blanks
        //     string[] blanks = new string[1];
        //     blanks[0] = "__________________";
        //     List<XElement> blankList = El.NumberList(blanks);


        //     foreach (string word in shuffledVocab.Keys)
        //     {
        //         // Definition
        //         XElement definitionParagraph = El.Paragraph(vocab[word]);
        //         El.SetParagraphSpacing(definitionParagraph, 0, 0);
        //         definitionParagraph.Add(El.InlineBreak());

        //         // Add table row
        //         List<List<XElement>> content = [
        //             blankList,
        //             [definitionParagraph]
        //         ];
        //         mainTable.Add(El.TableRow(tableColWidths, content));
        //     }

        //     // Answer list
        //     foreach (XElement item in El.NumberList(shuffledVocab.Keys))
        //         answerList.Add(item);

        //     return (mainTable, answerList);
        // }

        public static (List<OpenXmlElement>, List<OpenXmlElement>) GetProcessedVocab(OpenXmlElementList allElements, int sectionNo = -1)
        {
            List<OpenXmlElement> elements = GetParagraphsByIdentifier(allElements, "VOCAB");
            foreach (OpenXmlElement element in elements)
                Console.WriteLine(element.OuterXml);
            List<OpenXmlElement> mainActivity = [];
            List<OpenXmlElement> answerKey = [];

            // Format & add title
            mainActivity.Add(GetFormattedSectionTitleElement("Vocabulary", sectionNo));

            // Get vocab
            Dictionary<string, string> vocab = GetVocab(elements);

            // Vocab box
            mainActivity.Add(VocabBox(vocab.Keys));
            mainActivity.Add(new Paragraph());

            // Blanks and definitions
            // (XElement blanksAndDefinitions, List<XElement> answerList) = VocabBlanksAndDefinitions(vocab);
            // mainActivity.Append(blanksAndDefinitions);

            // Page break
            // mainActivity.Append(El.PageBreak());

            // Answer key
            // answerKey.Append(GetFormattedAnswerKeySectionTitleElement("Vocabulary", sectionNo));
            // foreach (XElement paragraph in answerList)
            //     answerKey.Append(paragraph);

            return (mainActivity, answerKey);
        }

        // public static List<XElement> GetProcessedReading(IEnumerable<XElement> allParagraphs, int sectionNo = -1)
        // {
        //     List<XElement> paragraphs = GetParagraphsByIdentifier(allParagraphs, "READING");
        //     List<XElement> result = [];
        //     XElement? origTitleElement = GetSectionTitleElement(paragraphs);

        //     // Format & add title
        //     if (origTitleElement != null)
        //     {
        //         XElement newTitleElement = GetFormattedSectionTitleElement(
        //             RemovePrefix((string)origTitleElement),
        //             sectionNo
        //         );
        //         result.Add(newTitleElement);
        //     }

        //     // Filter paragraphs
        //     foreach (XElement paragraph in paragraphs)
        //     {
        //         if (!StartsWith(paragraph, "title:"))
        //         {
        //             // Add styling
        //             if (IsImage(paragraph))
        //                 FormatImage(paragraph);
        //             else
        //             {
        //                 El.SetDefaultParagraphStyles(paragraph);
        //                 El.SetParagraphLine(paragraph, 276);
        //             }

        //             result.Add(paragraph);
        //         }
        //     }

        //     // Page break
        //     result.Add(El.PageBreak());

        //     return result;
        // }

        // public static (List<XElement>, List<XElement>) TrueOrFalseQs(List<XElement> paragraphs)
        // {
        //     List<XElement> mainActivity = [];

        //     // True-or-False header
        //     XElement header = El.Paragraph("Circle \"T\" for True or \"F\" for False for the following statements:");
        //     El.AddBoldToParagraph(header);
        //     mainActivity.Add(header);

        //     Dictionary<string, string> TFStatements = [];
        //     for (int i = 0; i < paragraphs.Count; i += 2)
        //     {
        //         if (StartsWith(paragraphs[i + 1], "t"))
        //             TFStatements.Add((string)paragraphs[i], "T");
        //         else
        //             TFStatements.Add((string)paragraphs[i], "F");
        //     }

        //     // Create table
        //     int[] tableColWidths = [9639, 851, 662];
        //     XElement mainTable = El.Table(
        //         tableColWidths,
        //         El.TableBorderAttributes("none", 0, 0, "auto")
        //     );

        //     // Turn statements into a list
        //     List<XElement> TFStatementList = El.NumberList(TFStatements.Keys);

        //     // Add statements to table
        //     foreach (XElement item in TFStatementList)
        //     {
        //         mainTable.Add(El.TableRow(tableColWidths, [
        //             [item],
        //             [El.Paragraph("T")],
        //             [El.Paragraph("F")]
        //         ]));
        //     }

        //     // Add table
        //     mainActivity.Add(mainTable);

        //     // Answer list
        //     List<XElement> answerList = El.NumberList(TFStatements.Values);

        //     return (mainActivity, answerList);
        // }

        // public static (List<XElement>, List<XElement>) GetProcessedCompQs(IEnumerable<XElement> allParagraphs, int sectionNo = -1)
        // {
        //     List<XElement> mainActivity = [];
        //     List<XElement> answerKey = [];

        //     if (ContainsIdentifier(allParagraphs, "QUESTIONS"))
        //     {
        //         // Format & add title
        //         mainActivity.Add(GetFormattedSectionTitleElement("Vocabulary", sectionNo));

        //         // True-or-False questions
        //         List<XElement> TFParagraphs = GetParagraphsByIdentifier(allParagraphs, "TF");
        //         (List<XElement> TFQuestions, List<XElement> TFAnswerKey) = TrueOrFalseQs(TFParagraphs);
        //         foreach (XElement paragraph in TFQuestions)
        //             mainActivity.Add(paragraph);

        //         // Multiple choice questions
        //         List<XElement> MCParagraphs = GetParagraphsByIdentifier(allParagraphs, "MC");

        //         // Page break
        //         mainActivity.Add(El.PageBreak());

        //         // Answer key
        //         answerKey.Add(GetFormattedAnswerKeySectionTitleElement("Vocabulary", sectionNo));
        //         foreach (XElement paragraph in TFAnswerKey)
        //             answerKey.Add(paragraph);
        //         answerKey.Add(El.Paragraph());
        //     }

        //     return (mainActivity, answerKey);
        // }

        // public static XElement AnswerKeyTitleElement()
        // {
        //     XElement titleElement = El.Paragraph("ANSWER KEY");

        //     El.AddBoldToParagraph(titleElement);
        //     El.CenterParagraph(titleElement);
        //     El.SetParagraphSize(titleElement, 40);
        //     El.SetParagraphSpacing(titleElement, 0, 200);

        //     return titleElement;
        // }
    }
}