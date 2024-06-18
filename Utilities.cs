using System;
using System.Xml.Linq;
using WorksheetGenerator.Elements;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.ExtendedProperties;
using D = DocumentFormat.OpenXml.Drawing;

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
                if (ElementTextStartsWith(element, "title:"))
                {
                    string title = RemovePrefix(GetParagraphText(element)).ToUpper();

                    Paragraph worksheetTitlePara = new(
                        El.ParagraphStyle("WorksheetTitle"),
                        new Run(new Text(title))
                    );

                    return worksheetTitlePara;
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

        public static bool IsImage(OpenXmlElement element)
        {
            return element.Descendants<D.Blip>().Any();
        }

        public static List<Paragraph> GetParagraphsByIdentifier(OpenXmlElementList elements, string identifierName)
        {
            bool isBetweenIdentifiers = false;
            List<Paragraph> result = [];

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
                    if (!ElementTextStartsWith(element, "chatgpt:") && (HasText(element) || IsImage(element)))
                        result.Add((Paragraph)element);
            }

            return result;
        }

        public static Int64Value GetWidth(Int64Value width, Int64Value height, Int64Value desiredHeight)
        {
            double w = width;
            double h = height;
            double dH = desiredHeight;

            return (Int64Value)Math.Round((double)(dH / h * w));
        }

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

        public static Paragraph? GetSectionTitleElement(List<Paragraph> paragraphs)
        {
            foreach (Paragraph paragraph in paragraphs)
                if (ElementTextStartsWith(paragraph, "title:"))
                    return paragraph;

            return null;
        }

        public static Paragraph GetFormattedSectionTitleElement(string title, int sectionNo = -1)
        {
            string formattedTitle;
            if (sectionNo >= 0)
                formattedTitle = sectionNo + ". " + title.Trim();
            else
                formattedTitle = title.Trim();

            Paragraph titleElement = new(
                El.ParagraphStyle("SectionTitle"),
                new Run(new Text(formattedTitle))
            );

            return titleElement;
        }

        public static Paragraph GetFormattedAnswerKeySectionTitleElement(string title, int sectionNo = -1)
        {
            string formattedTitle;
            if (sectionNo >= 0)
                formattedTitle = sectionNo + ". " + title.Trim().ToUpper();
            else
                formattedTitle = title.Trim().ToUpper();

            return new Paragraph(
                El.ParagraphStyle("AnswerKeySectionTitle"),
                new Run(new Text(formattedTitle))
            );
        }

        public static Dictionary<string, string> GetVocab(List<Paragraph> elements)
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
                    if (definition[^1] == '.')
                        definition = definition[..^1];
                    vocab.Add(word, definition);
                }
            }

            return vocab;
        }

        public static Table VocabBox(ICollection<string> words)
        {
            string formattedWords = string.Join("        ", words);
            Paragraph formattedWordsPara = new(
                El.ParagraphStyle("VocabBox"),
                new Run(new Text(formattedWords))
            );

            Table box = new(
                El.TableStyle("Box"),
                new TableRow(
                    new TableCell(
                        new TableCellProperties(
                            El.TableCellMargin(440, 440, 0, 440),
                            El.TableCellWidth(11169)
                        ),
                        formattedWordsPara
                    )
                )
            );

            return box;
        }

        public static Dictionary<TKey, TValue> ShuffledDictionary<TKey, TValue>(Dictionary<TKey, TValue> dict) where TKey : notnull
        {
            Random random = new();

            // Convert the dictionary to a list of key-value pairs
            List<KeyValuePair<TKey, TValue>> keyValuePairs = [.. dict];

            // Shuffle the list using the Fisher-Yates algorithm
            int n = keyValuePairs.Count;
            while (n > 1)
            {
                n--;
                int k = random.Next(n + 1);
                (keyValuePairs[n], keyValuePairs[k]) = (keyValuePairs[k], keyValuePairs[n]);
            }

            // Create a new dictionary from the shuffled list
            return keyValuePairs.ToDictionary(pair => pair.Key, pair => pair.Value);
        }

        public static (Table, List<Paragraph>) VocabBlanksAndDefinitions(Dictionary<string, string> vocab)
        {
            int[] tableColWidths = [3261, 7938];

            // Shuffle words
            Dictionary<string, string> shuffledVocab = ShuffledDictionary(vocab);

            // Create main table
            Table mainTable = new(El.TableStyle("NoBorderTable"));

            // Blanks
            string[] blank = ["__________________"];
            List<Paragraph> blankList = El.NumberList(blank);

            // Add table rows
            foreach (string word in shuffledVocab.Keys)
            {
                mainTable.AppendChild(
                    new TableRow(
                        // Blank
                        new TableCell(
                            new TableCellProperties(
                                El.TableCellWidth(3261)
                            ),
                            blankList[0].CloneNode(true)
                        ),
                        new TableCell(
                            new TableCellProperties(
                                El.TableCellWidth(7938)
                            ),
                            new Paragraph(
                                El.ParagraphStyle("ListActivity"),
                                new Run(new Text(vocab[word]))
                            )
                        )
                    )
                );
            }

            // Create answer list
            List<Paragraph> answerList = El.NumberList(shuffledVocab.Keys);

            return (mainTable, answerList);
        }

        public static (List<OpenXmlElement>, List<Paragraph>) GetProcessedVocab(OpenXmlElementList allElements, int sectionNo = -1)
        {
            List<Paragraph> paragraphs = GetParagraphsByIdentifier(allElements, "VOCAB");
            List<OpenXmlElement> mainActivity = [];
            List<Paragraph> answerKey = [];

            // Format & add title
            mainActivity.Add(GetFormattedSectionTitleElement("Vocabulary", sectionNo));

            // Get vocab
            Dictionary<string, string> vocab = GetVocab(paragraphs);

            // Vocab box
            mainActivity.Add(VocabBox(vocab.Keys));
            mainActivity.Add(new Paragraph());

            // Blanks and definitions
            (Table blanksAndDefinitions, List<Paragraph> answerList) = VocabBlanksAndDefinitions(vocab);
            mainActivity.Add(blanksAndDefinitions);

            // Page break
            mainActivity.Add(El.PageBreak());

            // Answer key
            answerKey.Add(GetFormattedAnswerKeySectionTitleElement("Vocabulary", sectionNo));
            foreach (Paragraph paragraph in answerList)
                answerKey.Add(paragraph);

            return (mainActivity, answerKey);
        }

        public static List<Paragraph> GetProcessedReading(OpenXmlElementList allElements, Dictionary<string, string> imageRelIds, int sectionNo = -1)
        {
            List<Paragraph> paragraphs = GetParagraphsByIdentifier(allElements, "READING");
            // foreach (Paragraph paragraph in paragraphs)
            //     Console.WriteLine(paragraph.OuterXml);
            List<Paragraph> result = [];

            // Format & add title
            Paragraph? origTitleElement = GetSectionTitleElement(paragraphs);
            if (origTitleElement != null)
            {
                Paragraph newTitleElement = GetFormattedSectionTitleElement(
                    RemovePrefix(GetParagraphText(origTitleElement)),
                    sectionNo
                );
                result.Add(newTitleElement);
            }

            // Filter paragraphs
            foreach (Paragraph paragraph in paragraphs)
            {
                if (!ElementTextStartsWith(paragraph, "title:"))
                {
                    // Format images
                    if (IsImage(paragraph))
                    {
                        string? oldImageRelId = El.GetImageRelId(paragraph);
                        if (oldImageRelId != null)
                        {
                            Paragraph? image = El.Image(paragraph, imageRelIds[oldImageRelId], 2230120L, "InlineImage");
                            if (image != null)
                                result.Add(image);
                        }

                        // // Edit <a:prstGeom> (Rounded corners)
                        // XElement? geometryElement = element.Descendants(El.a + "prstGeom").FirstOrDefault();
                        // geometryElement?.SetAttributeValue("prst", "roundRect");
                    }
                    else
                        result.Add(new Paragraph(
                            El.ParagraphStyle("Paragraph"),
                            new Run(new Text(GetParagraphText(paragraph)))
                        ));
                    //     }
                }
            }

            // Page break
            result.Add(El.PageBreak());

            return result;
        }

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