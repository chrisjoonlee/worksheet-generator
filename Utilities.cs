using WorksheetGenerator.Elements;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using D = DocumentFormat.OpenXml.Drawing;

namespace WorksheetGenerator.Utilities
{
    public static class HF
    {
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

        public static bool ContainsIdentifier(OpenXmlElementList elements, string identifierName)
        {
            foreach (OpenXmlElement element in elements)
                if (IsIdentifier(element) && ElementTextStartsWith(element, identifierName))
                    return true;
            return false;
        }

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
                    if (!ElementTextStartsWith(element, "chatgpt:") && element is not SectionProperties)
                        result.Add((Paragraph)element);
            }

            return result;
        }

        public static List<Paragraph> NoEmptyParagraphs(List<Paragraph> paragraphs)
        {
            List<Paragraph> result = [];

            foreach (Paragraph paragraph in paragraphs)
                if (HasText(paragraph) || IsImage(paragraph))
                    result.Add(paragraph);

            return result;
        }

        public static Int64Value GetWidth(Int64Value width, Int64Value height, Int64Value desiredHeight)
        {
            double w = width;
            double h = height;
            double dH = desiredHeight;

            return (Int64Value)Math.Round((double)(dH / h * w));
        }

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
            List<Paragraph> paragraphs = NoEmptyParagraphs(GetParagraphsByIdentifier(allElements, "VOCAB"));
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
            List<Paragraph> paragraphs = NoEmptyParagraphs(GetParagraphsByIdentifier(allElements, "READING"));
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

        public static (List<OpenXmlElement>, List<Paragraph>) TrueOrFalseQs(List<Paragraph> paragraphs)
        {
            List<OpenXmlElement> mainActivity = [];

            // True-or-False header
            Paragraph header = new(
                El.ParagraphStyle("SubsectionTitle"),
                new Run(new Text("Circle \"T\" for True or \"F\" for False for the following statements:")
            ));
            mainActivity.Add(header);

            // Organize statements and answers
            Dictionary<string, string> TFStatements = [];
            for (int i = 0; i < paragraphs.Count; i += 2)
            {
                if (ElementTextStartsWith(paragraphs[i + 1], "t"))
                    TFStatements.Add(GetParagraphText(paragraphs[i]), "T");
                else
                    TFStatements.Add(GetParagraphText(paragraphs[i]), "F");
            }

            // Turn statements into a list
            List<Paragraph> TFStatementList = El.NumberList(TFStatements.Keys, "ListActivity");

            // Add statements to table
            Table mainTable = new(El.TableStyle("NoBorderTable"));
            foreach (Paragraph item in TFStatementList)
            {
                mainTable.AppendChild(
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(
                                El.TableCellWidth(9639)
                            ),
                            item
                        ),
                        new TableCell(
                            new TableCellProperties(
                                El.TableCellWidth(851)
                            ),
                            new Paragraph(
                                El.ParagraphStyle("Text"),
                                new Run(new Text("T"))
                            )
                        ),
                        new TableCell(
                            new TableCellProperties(
                                El.TableCellWidth(662)
                            ),
                            new Paragraph(
                                El.ParagraphStyle("Text"),
                                new Run(new Text("F"))
                            )
                        )
                    )
                );
            }

            // Add table
            mainActivity.Add(mainTable);

            // Answer list
            List<Paragraph> answerList = El.NumberList(TFStatements.Values);

            return (mainActivity, answerList);
        }

        public static List<List<Paragraph>> GetParagraphChunks(List<Paragraph> paragraphs)
        {
            List<List<Paragraph>> chunkedList = new();
            List<Paragraph> currentChunk = new();

            foreach (Paragraph paragraph in paragraphs)
            {
                if (HasText(paragraph))
                    currentChunk.Add(paragraph);
                else
                {
                    if (currentChunk.Count > 0)
                    {
                        chunkedList.Add(currentChunk);
                        currentChunk = new();
                    }
                }
            }

            if (currentChunk.Count > 0)
                chunkedList.Add(currentChunk);

            return chunkedList;
        }

        public static bool HasBold(Paragraph paragraph)
        {
            return paragraph.Descendants<Run>().Any(run => IsBoldRun(run));
        }

        public static bool IsBoldRun(Run run)
        {
            return run.RunProperties != null && run.RunProperties.Bold != null;
        }

        public static T[] ShuffledArray<T>(T[] array)
        {
            T[] shuffledArray = new T[array.Length];
            Array.Copy(array, shuffledArray, array.Length);

            Random random = new Random();

            // Perform Fisher-Yates shuffle
            for (int i = shuffledArray.Length - 1; i > 0; i--)
            {
                int j = random.Next(0, i + 1);
                (shuffledArray[j], shuffledArray[i]) = (shuffledArray[i], shuffledArray[j]);
            }

            return shuffledArray;
        }

        public static (List<Paragraph>, List<Paragraph>) MultipleChoiceQs(List<Paragraph> paragraphs)
        {
            List<Paragraph> mainActivity = [];
            List<Paragraph> answerList = [];

            // Multiple choice header
            Paragraph header = new(
                El.ParagraphStyle("SubsectionTitle"),
                new Run(new Text("Choose the correct answer for each of the following questions:")
            ));
            mainActivity.Add(header);

            // Organize multiple choice questions
            List<MultipleChoice> multipleChoiceQs = [];
            List<List<Paragraph>> paragraphChunks = GetParagraphChunks(paragraphs);
            foreach (List<Paragraph> paragraphChunk in paragraphChunks)
            {
                // Question
                string question = GetParagraphText(paragraphChunk[0]);
                string[] choices = new string[paragraphChunk.Count - 1];
                string? answer = null;

                for (int i = 1; i < paragraphChunk.Count; i++)
                {
                    choices[i - 1] = GetParagraphText(paragraphChunk[i]);
                    if (HasBold(paragraphChunk[i]))
                        answer = GetParagraphText(paragraphChunk[i]);
                }

                if (answer != null)
                    multipleChoiceQs.Add(new MultipleChoice(question, choices, answer));
            }

            int q_no = 1;
            foreach (MultipleChoice mc in multipleChoiceQs)
            {
                // Shuffle choices
                if (!mc.Choices[^1].Equals("all of the above", StringComparison.CurrentCultureIgnoreCase))
                    mc.Choices = ShuffledArray(mc.Choices);

                // Format text
                mainActivity.Add(new Paragraph(
                    El.ParagraphStyle("Text"),
                    new Run(new Text($"{q_no}. {mc.Question}"))
                ));
                mainActivity.Add(new Paragraph(El.ParagraphStyle("Text")));

                // Keep track of question number
                q_no++;
            }

            return (mainActivity, answerList);
        }

        public static (List<OpenXmlElement>, List<Paragraph>) GetProcessedCompQs(OpenXmlElementList allElements, int sectionNo = -1)
        {
            List<OpenXmlElement> mainActivity = [];
            List<Paragraph> answerKey = [];

            if (ContainsIdentifier(allElements, "QUESTIONS"))
            {
                // Format & add title
                mainActivity.Add(GetFormattedSectionTitleElement("Comprehension Questions", sectionNo));

                // True-or-False questions
                List<Paragraph> TFParagraphs = NoEmptyParagraphs(GetParagraphsByIdentifier(allElements, "TF"));
                (List<OpenXmlElement> TFQuestions, List<Paragraph> TFAnswerKey) = TrueOrFalseQs(TFParagraphs);
                foreach (OpenXmlElement element in TFQuestions)
                    mainActivity.Add(element);

                // Multiple choice questions
                List<Paragraph> MCParagraphs = GetParagraphsByIdentifier(allElements, "MC");
                (List<Paragraph> MCQuestions, List<Paragraph> MCAnswerKey) = MultipleChoiceQs(MCParagraphs);
                foreach (OpenXmlElement paragraph in MCQuestions)
                    mainActivity.Add(paragraph);

                // Page break
                mainActivity.Add(El.PageBreak());

                // Answer key
                answerKey.Add(GetFormattedAnswerKeySectionTitleElement("Vocabulary", sectionNo));
                foreach (Paragraph paragraph in TFAnswerKey)
                    answerKey.Add(paragraph);
                answerKey.Add(new Paragraph());
                foreach (Paragraph paragraph in MCAnswerKey)
                    answerKey.Add(paragraph);
            }

            return (mainActivity, answerKey);
        }

        public static Paragraph AnswerKeyTitleElement()
        {
            Paragraph titlePara = new Paragraph(
                El.ParagraphStyle("AnswerKeyTitle"),
                new Run(new Text("ANSWER KEY"))
            );

            return titlePara;
        }
    }
}