using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using OpenMcdf;

namespace AltiumTranslate
{
    class Program
    {
        // Reversing options
        static bool _DoNotCopyKnownKeys = false;
        static bool _CopyUnknownKeys = false;

        enum inputType
        {
            AD,
            CS
        };

        static inputType inputFileType;
        static Dictionary<string, string> CsToAdDictionary;
        static CompoundFile PCBDOC_InputFile, PCBDOC_OutputFile;
        static string inputFileName;
        static byte[] AD_FileHeaderSix = { 0x13, 0x00, 0x00, 0x00, 0x13, 0x50, 0x43, 0x42, 0x20, 0x36, 0x2E, 0x30, 0x20, 0x42, 0x69, 0x6E, 0x61, 0x72, 0x79, 0x20, 0x46, 0x69, 0x6C, 0x65, 0x0A, 0xD7, 0xA3, 0x70, 0x3D, 0x0A, 0x14, 0x40 };
        static byte[] CS_FileHeaderSix = { 0x21, 0x00, 0x00, 0x00, 0x21, 0x43, 0x69, 0x72, 0x63, 0x75, 0x69, 0x74, 0x53, 0x74, 0x75, 0x64, 0x69, 0x6F, 0x20, 0x50, 0x43, 0x42, 0x20, 0x36, 0x2E, 0x30, 0x20, 0x42, 0x69, 0x6E, 0x61, 0x72, 0x79, 0x20, 0x46, 0x69, 0x6C, 0x65, 0x0A, 0xD7, 0xA3, 0x70, 0x3D, 0x0A, 0x14, 0x40 };

        private static void ReadTranslationList()
        {
            CsToAdDictionary = new Dictionary<string, string>();
            var lines = File.ReadAllLines(Path.Combine(Environment.CurrentDirectory, "translate.csv"));
            foreach(var line in lines)
            {
                var values = line.Split(',');
                if (values.Length == 2)
                {
                    CsToAdDictionary.Add(values[0].Trim(), values[1].Trim());
                } else
                {
                    Console.WriteLine("Malformed key-value pair found - " + line);
                }
            }
        }

        private static void CopyStorage(string from, string to)
        {
            var fromStore = PCBDOC_InputFile.RootStorage.GetStorage(from);
            var toStore = PCBDOC_OutputFile.RootStorage.AddStorage(to);

            Action<CFItem> va = delegate (CFItem item)
            {
                if (item.IsStream)
                {
                    var toStream = toStore.AddStream(item.Name);
                    var fromStream = fromStore.GetStream(item.Name);
                    toStream.SetData(fromStream.GetData());
                    //if(toStream.Size > 4)
                    //{
                    //}
                }
            };

            fromStore.VisitEntries(va, false);
        }

        private static void ADInputVisitEntries()
        {
            Action<CFItem> storageActionDelegate = delegate (CFItem item)
            {
                if (item.IsStorage)
                {

                    if (CsToAdDictionary.ContainsValue(item.Name))
                    {
                        if (!_DoNotCopyKnownKeys)
                            CopyStorage(item.Name, CsToAdDictionary.FirstOrDefault(x => x.Value == item.Name).Key);
                    }
                    else if (_CopyUnknownKeys)
                    {
                        CopyStorage(item.Name, item.Name);
                    }

                }
            };

            PCBDOC_InputFile.RootStorage.VisitEntries(storageActionDelegate, false);
        }

        private static void CSInputVisitEntries()
        {
            Action<CFItem> storageActionDelegate = delegate (CFItem item)
            {
                if (item.IsStorage)
                {

                    if (CsToAdDictionary.ContainsKey(item.Name))
                    {
                        if (!_DoNotCopyKnownKeys)
                            CopyStorage(item.Name, CsToAdDictionary[item.Name]);
                    }
                    else if (_CopyUnknownKeys)
                    {
                        CopyStorage(item.Name, item.Name);
                    }

                }
            };

            PCBDOC_InputFile.RootStorage.VisitEntries(storageActionDelegate, false);
        }

        private static void LoadNewPCBFile()
        {
            if (inputFileType == inputType.CS)
            {
                PCBDOC_OutputFile.RootStorage.AddStream("FileHeaderSix").SetData(AD_FileHeaderSix);
            } else
            {
                PCBDOC_OutputFile.RootStorage.AddStream("FileHeaderSix").SetData(CS_FileHeaderSix);
            }

            PCBDOC_OutputFile.RootStorage.AddStream("FileHeader").SetData(
                PCBDOC_InputFile.RootStorage.GetStream("FileHeader").GetData());


            if(inputFileType == inputType.CS)
            {
                CSInputVisitEntries();
            }
            else
            {
                ADInputVisitEntries();
            }

        }

        private static string GetNewFileName()
        {
            var extention = inputFileType == inputType.AD ? ".CSPcbDoc" : ".PcbDoc"; 
            var newFileName = Path.GetFileNameWithoutExtension(inputFileName) + extention;
            var dir = Path.GetDirectoryName(inputFileName);
            return Path.Combine(dir, newFileName);
        }

        static void Main(string[] args)
        {
            if(args.Length != 1)
            {
                Console.WriteLine("This takes a single argument of the input PCB filename");
                return;
            }

            if(!File.Exists(args[0]))
            {
                Console.WriteLine("Could not find input file: " + args[0]);
                return;
            }


            if(!File.Exists(Path.Combine(Environment.CurrentDirectory, "translate.csv")))
            {
                Console.WriteLine("Could not find translate.csv file, make sure it's in the executing directory.");
                return;
            }

            inputFileName = args[0];
            var inputFileExtention = Path.GetExtension(inputFileName).ToUpper();
            if (inputFileExtention == ".CSPCBDOC") {
                inputFileType = inputType.CS;
            }
            else if (inputFileExtention == ".PCBDOC")
            {
                inputFileType = inputType.AD;
            }
            else
            {
                Console.WriteLine("Unknown input file extention");
                return;
            }

            ReadTranslationList();

            PCBDOC_InputFile = new CompoundFile(inputFileName, CFSUpdateMode.ReadOnly, CFSConfiguration.Default);
            PCBDOC_OutputFile = new CompoundFile();

            LoadNewPCBFile();
            PCBDOC_OutputFile.Save(GetNewFileName());
            

            Console.WriteLine("Translation complete.");
        }


    }
}
