using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.IO;

namespace DisciplineWorkProgram.Word.Helpers
{
    public static class Ooxml
    {
        public static IEnumerable<T> FindElementsByBookmark<T>(BookmarkStart bookmarkStart, uint outerLevels) where T : OpenXmlElement
        {
            var elements = new List<T>();
            var elem = bookmarkStart.NextSibling();

            while (elem != null)
            {
                //Проверка самого элемента
                switch (elem)
                {
                    case BookmarkEnd el when el.Id == bookmarkStart.Id:
                        return elements;
                    case T element:
                        elements.Add(element);
                        break;
                }
                //Проверка всех элементов "под" самим элементом
                foreach (var node in elem.Descendants())
                {
                    switch (node)
                    {
                        case BookmarkEnd end when end.Id == bookmarkStart.Id:
                            return elements;
                        case T n:
                            elements.Add(n);
                            continue;
                    }
                }

                var next = elem.NextSibling();

                if (!(next is null))
                    elem = next;
                else if (outerLevels > 0)
                {
                    elem = elem.Parent;
                    outerLevels--;
                }
                else return elements;
            }

            return elements;
        }

        public static IDictionary<string, BookmarkStart> GetBookmarks(WordprocessingDocument doc, string bookmarkStartName)
        {
            var bookmarkMap = new Dictionary<string, BookmarkStart>();

            foreach (var bookmarkStart in doc.MainDocumentPart.Document.Body.Descendants<BookmarkStart>())
                if (bookmarkStart.Name.ToString().StartsWith(bookmarkStartName)) //Если сделать без идентификатора, то будут коллизии
                    bookmarkMap[bookmarkStart.Name.ToString().Substring(8)] = bookmarkStart; //Но Autofill отсекается

            return bookmarkMap;
        }

        public static void SaveDoc(OpenXmlPackage doc, string dir, string name)
        {
            //var regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            name = RegexPatterns.InvalidChars.Replace(name, "");

            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            doc.SaveAs($"{dir}/{name}.docx");
        }
    }
}