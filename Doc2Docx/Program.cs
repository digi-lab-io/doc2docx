/*
Copyright 2015 Softcom

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

   http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

**/

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
