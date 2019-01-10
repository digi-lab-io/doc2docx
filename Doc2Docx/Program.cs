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
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace Doc2Docx
{
    class Program
    {
        private int _processId;
        Word.Application _word;

        static void Main(string[] args)
        {
            var filesToConvert = new List<string>();
            var so = SearchOption.TopDirectoryOnly;
            if(args.Length > 0 && (args[0] == "/r" || args[0] == "-r" || args[0] == "--recursive"))
            {
                args = args.Skip(1).ToArray();
                so = SearchOption.AllDirectories;
            }
            foreach (var arg in args)
            {
                var ext = Path.GetExtension(arg);
                if (ext != ".doc" && ext != ".xml" && ext != ".wml")
                {
                    Console.WriteLine($"Skipping {arg}: only DOC / WML / XML files permitted.");
                    continue;
                }
                if (arg.Contains('*'))
                {
                    filesToConvert.AddRange(Directory.GetFiles(".", arg, so).Select(_ => Path.GetFullPath(_)));
                }
                else
                {
                    filesToConvert.Add(Path.GetFullPath(arg));
                }
                Console.WriteLine($"Accepted {arg}");
            }
            if (filesToConvert.Count > 0)
            {
                filesToConvert = filesToConvert.Distinct().ToList();
                new Program(filesToConvert);
            }
        }

        internal Program(IEnumerable<string> filesToConvert)
        {
            StartWord();
            try
            {
                foreach (var file in filesToConvert)
                {
                    try
                    {
                        var path = Path.GetFullPath(file);
                        ConvertFile(path);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"{file}: {e.Message}");
                    }
                }
            }
            finally
            {
                ShutdownWord();
            }

        }

        private void ConvertFile(string filepath)
        {
            var fileformat = Word.WdSaveFormat.wdFormatXMLDocument;

            var newfilepath = Path.ChangeExtension(filepath, ".docx");
            Console.WriteLine($"{Path.GetFileName(filepath)} -> {Path.GetFileName(newfilepath)}");
            Word._Document document = _word.Documents.Open(
                filepath,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                PasswordDocument: Type.Missing,
                PasswordTemplate: Type.Missing,
                Revert: false,
                WritePasswordDocument: Type.Missing,
                WritePasswordTemplate: Type.Missing,
                Format: Type.Missing,
                Encoding: Type.Missing,
                Visible: false,
                OpenAndRepair: false,
                DocumentDirection: Type.Missing,
                NoEncodingDialog: true,
                XMLTransform: Type.Missing);
            document.Convert();
            string v = _word.Version;
            switch (v)
            {
                case "14.0":
                case "15.0":
                    document.SaveAs2(newfilepath, fileformat, CompatibilityMode: Word.WdCompatibilityMode.wdWord2007);
                    break;
                default:
                    document.SaveAs(newfilepath, fileformat);
                    break;
            }
            document.Close();
            document = null;
        }


        void StartWord()
        {
            string wordAppId = "" + DateTime.Now.Ticks;
            _word = new Word.Application();
            _word.Application.Caption = wordAppId;
            _word.Application.Visible = true;
            _processId = GetProcessIdByWindowTitle(wordAppId);
            _word.Application.Visible = false;
        }

        private void ShutdownWord()
        {
            if (_word != null)
            {
                _word.Quit();
                _word = null;
                Process process = Process.GetProcessById(_processId);
                process.Kill();
            }
        }

        public static int GetProcessIdByWindowTitle(string paramWordAppId)
        {
            Process[] P_CESSES = Process.GetProcessesByName("WINWORD");
            for (int p_count = 0; p_count < P_CESSES.Length; p_count++)
            {
                if (P_CESSES[p_count].MainWindowTitle.Equals(paramWordAppId))
                {
                    return P_CESSES[p_count].Id;
                }
            }
            return Int32.MaxValue;
        }

    }
}
