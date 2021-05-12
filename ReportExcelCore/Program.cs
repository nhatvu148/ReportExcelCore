using Microsoft.VisualBasic.FileIO;
using OfficeOpenXml;
using System;
using System.IO;
using System.Reflection;

namespace ReportExcelCore
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string pathDirectory = $"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}";
            var fileSource = new FileInfo($"{pathDirectory}/../input/DLP-Prediction.xlsm");
            var fileDestination = new FileInfo($"{pathDirectory}/../output/DLP-Prediction1.xlsm");

            int alpp = 0;

            using (TextFieldParser parser = new TextFieldParser($"{pathDirectory}/../input/SHIPINP.DAT"))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                int counter = 0;
                while (!parser.EndOfData)
                {
                    //Process row
                    string[] fields = parser.ReadFields();

                    foreach (string field in fields)
                    {
                        if (counter++ == 4)
                        {
                            //TODO: Process field
                            string[] ssize = field.Split(null);
                            alpp = Int32.Parse(ssize[0]);
                            break;
                        }
                    }
                }
            }

            using (TextFieldParser parser = new TextFieldParser($"{pathDirectory}/../input/INPPARA.DAT"))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                int counter = 0;
                while (!parser.EndOfData)
                {
                    //Process row
                    string[] fields = parser.ReadFields();

                    foreach (string field in fields)
                    {
                        if (counter++ == 4)
                        {
                            //TODO: Process field
                            string[] ssize = field.Split(null);
                            alpp = Int32.Parse(ssize[0]);
                            break;
                        }
                    }
                }
            }


            using (var excelFileSource = new ExcelPackage(fileSource))
            using (var excelFileDestination = new ExcelPackage(fileDestination))
            {
                var worksheetSource = excelFileSource.Workbook.Worksheets[1];
                worksheetSource.Cells[5, 5].Value = alpp;
                excelFileSource.SaveAs(fileDestination);
            }

            Console.WriteLine("Hello World!");
        }
    }
}
