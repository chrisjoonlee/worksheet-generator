using System.Xml.Linq;

class AddToDoc
{
    public static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("Usage: dotnet run <folder_name>");
            return;
        }

        // Load document
        string filePath = $"../../{args[0]}/word/document.xml";
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XDocument doc = XDocument.Load(filePath);

        XElement? root = doc.Element(w + "document");
        if (root != null)
        {
            XElement? body = root.Element(w + "body");
            if (body != null)
            {

                string? answer = "";
                while (answer != "1" && answer != "2")
                {
                    Console.Write("Add to beginning (1) or end (2) of doc? ");
                    answer = Console.ReadLine();
                }

                // Ask user for text to add
                Console.WriteLine("Enter text to add:");
                string? text = Console.ReadLine();
                XElement new_element = new XElement(w + "p",
                                            new XElement(w + "pPr",
                                                new XElement(w + "lang",
                                                    new XAttribute(w + "val", "en-US"))),
                                            new XElement(w + "r",
                                                new XElement(w + "rPr",
                                                    new XElement(w + "lang", "en-US")),
                                                new XElement(w + "t", text)));

                // Add text
                if (answer == "1")
                {
                    body.AddFirst(new_element);
                }
                else
                {
                    XElement? sectPr = body.Descendants(w + "sectPr").FirstOrDefault();
                    if (sectPr != null)
                    {
                        sectPr.AddBeforeSelf(new_element);
                    }
                }

                // Save
                doc.Save(filePath);
                Console.WriteLine("Changes saved.");
            }
        }
    }
}

