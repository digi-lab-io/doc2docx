using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;

namespace Doc2Docx
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {
                FileInfo file = new FileInfo(@args[0]);

                if (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".xml" || file.Extension.ToLower() == ".wml")
                {
                    Word._Application application = new Word.Application();
                    object fileformat = Word.WdSaveFormat.wdFormatXMLDocument;

                    object filename = file.FullName;
                    object newfilename = Path.ChangeExtension(file.FullName, ".docx");
                    Word._Document document = application.Documents.Open(filename);

                    document.Convert();
                    document.SaveAs(newfilename, fileformat);
                    document.Close();
                    document = null;

                    application.Quit();
                    application = null;
                }
            }
            catch (IndexOutOfRangeException e)
            {
                Console.WriteLine("Missing parameter: IndexOutOfRangeException {0}", e);
            }

        }
    }
}
