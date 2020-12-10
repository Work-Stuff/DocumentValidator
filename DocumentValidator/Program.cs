using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DocumentValidator
{
    class Program
    {
        static void Main(string[] args)
        {
            OpenXmlPackage wordDocument1 = WordprocessingDocument.Open("BouwstenenAlle.docm", false);
            List<ValidationErrorInfo> errorInfos1 = ValidateDocument(wordDocument1);

            foreach (ValidationErrorInfo info in errorInfos1)
            {
                Console.WriteLine(string.Format("BouwstenenAlle.docm Error Occurred {0}", info.Description));
            }

            OpenXmlPackage wordDocument2 = WordprocessingDocument.Open("LettertypenAlle.docx", false);
            List<ValidationErrorInfo> errorInfos2 = ValidateDocument(wordDocument2);

            foreach (ValidationErrorInfo info in errorInfos2)
            {
                Console.WriteLine(string.Format("LettertypenAlle.docx Validaton Error Occurred {0}", info.Description));
            }

            Console.ReadKey();
        }

        static List<ValidationErrorInfo> ValidateDocument(OpenXmlPackage package)
        {
            OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Office2016);
            List<ValidationErrorInfo> validationErrors = validator.Validate(package).ToList();

            return validationErrors;
        }
    }
}
