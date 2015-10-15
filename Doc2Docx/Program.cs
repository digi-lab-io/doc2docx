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
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace Doc2Docx
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo file = new FileInfo(@args[0]);

            if (file.Extension.ToLower() == ".doc" || file.Extension.ToLower() == ".xml" || file.Extension.ToLower() == ".wml")
            {

                Console.WriteLine("Starting document conversion...");
                // Get all running winword processes
                List<int> processIds = new List<int>();
                foreach (Process process in Process.GetProcessesByName("WINWORD"))
                {
                    processIds.Add(process.Id);
                }

                Word._Application word = new Word.Application();

                try
                {

                    object fileformat = Word.WdSaveFormat.wdFormatXMLDocument;

                    object filename = file.FullName;
                    object newfilename = Path.ChangeExtension(file.FullName, ".docx");
                    Word._Document document = word.Documents.Open(filename);
                    Console.WriteLine("Converting document '{0}' to DOCX.", file);

                    document.Convert();
                    document.SaveAs(newfilename, fileformat);
                    document.Close();
                    document = null;

                    word.Quit();
                    word = null;
                    Console.WriteLine("Success, quitting Word.");

                }
                catch (Exception e)
                {
                    Console.WriteLine("Error ocurred: {0}", e);
                }
                finally
                {

                    // and here is how it fails
                    foreach (Process process in Process.GetProcessesByName("WINWORD"))
                    {
                        if (!processIds.Contains(process.Id))
                        {
                            Console.WriteLine("Terminating Winword process with the ID: '{0}'.", process.Id);
                            process.Kill();
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("Only DOC / WML / XML files possible.");
            }

        }
    }
}
