using System;
using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Collections.Generic;
using System.Linq;

namespace OpenXmlPowerTools
{
    class ValidateFile
    {
        static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine("You must supply the file name/path of a DOCX file to validate.");
                Environment.Exit(0);
            }

            var fileName = args[0];
            FileInfo fi = new FileInfo(fileName);
            if (!fi.Exists)
            {
                Console.WriteLine("Error, {0} does not exist.", fileName);
                Environment.Exit(0);
            }
            if (Util.IsWordprocessingML(fi.Extension))
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(fileName, false))
                {
                    var fileFormatVersion = FileFormatVersions.Office2013;
                    OpenXmlValidator validator = new OpenXmlValidator(fileFormatVersion);
                    var errors = validator.Validate(wDoc);
                    bool valid = errors.Count() == 0;
                    if (valid)
                        Console.WriteLine("Document is valid.");
                    else
                    {
                        foreach (var errorInfo in errors)
                        {
                            Console.WriteLine(errorInfo.Description);
                            Console.WriteLine(" > " + errorInfo.Path.XPath);
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("Error, only works with DOCX files.");
            }
        }
    }
}
