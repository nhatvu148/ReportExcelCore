using Microsoft.VisualBasic.FileIO;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
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

            var inPredictionPath = $"{pathDirectory}/../input/DLP-Prediction.xlsm";
            var inRAOoutputPath = $"{pathDirectory}/../input/RAOoutput.csv";
            var outPredictionPath = $"{pathDirectory}/../output/DLP-Prediction1.xlsm";

            var shipinpPath = $"{pathDirectory}/../input/SHIPINP.DAT";
            var inpparaPath = $"{pathDirectory}/../input/INPPARA.DAT";

            if (args == null || args.Length == 0)
            {
                // no arguments
                //startTime = "2021-01-15 23:40:00";
                //endTime = "2021-01-25 23:00:00";
            }
            else
            {
                inPredictionPath = Convert.ToString(args[0]);
                inRAOoutputPath = Convert.ToString(args[1]);
                outPredictionPath = inPredictionPath;
                shipinpPath = Convert.ToString(args[2]);
                inpparaPath = Convert.ToString(args[3]);
            }

            var fileSource = new FileInfo(inPredictionPath);
            var fileDestination = new FileInfo(outPredictionPath);

            // Read data from RAOoutput.csv, Copy RAOoutput.csv to Data sheet
            using (TextFieldParser parser = new TextFieldParser(inRAOoutputPath))
            using (ExcelPackage excelFile = new ExcelPackage(fileSource))
            {
                var worksheet = excelFile.Workbook.Worksheets[0];
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                int row = 1;
                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    int col = 1;
                    foreach (string field in fields)
                    {
                        double number;

                        bool success = double.TryParse(field, out number);
                        if (success)
                        {
                            // Copy number to data sheet
                            // Check if number is an integer or double
                            worksheet.Cells[row, col++].Value = (number == (int)number) ? (int)number : number;
                        }
                        else
                        {
                            // Copy string to data sheet
                            worksheet.Cells[row, col++].Value = field;
                        }
                    }

                    row++;
                }

                excelFile.SaveAs(fileSource);
            }

            double alpp = 0;
            double vhw = 0;
            double speed = 0;
            int iramform = 0;
            int nchi = 0;
            double nram = 0;

            var waveH = new List<Dictionary<string, double>>();
            var waveDir = new List<Dictionary<string, double>>();
            var waveL = new List<Dictionary<string, double>>();

            // Read data from shipinp.dat
            using (TextFieldParser parser = new TextFieldParser(shipinpPath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                int counter = 0;
                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();

                    foreach (string field in fields)
                    {
                        if (counter++ == 4)
                        {
                            string[] ssize = field.Split(null);
                            // Assign first value of this row to alpp
                            alpp = double.Parse(ssize[0]);
                            break;
                        }
                    }
                }
            }

            // Read data from INPPARA.dat, 2 parsers for 2 reader iterations
            using (TextFieldParser parser = new TextFieldParser(inpparaPath))
            using (TextFieldParser parser2 = new TextFieldParser(inpparaPath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");

                int nhwIndex = 0, nchiIndex = 0, nramIndex = 0, outputParamIndex = 0, ipaccfouIndex = 0, nhaccoutIndex = 0,
                    iprsfouIndex = 0, idif2outIndex = 0;
                int idxCounter = 0;
                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();

                    // Extract indices for specified rows
                    foreach (string field in fields)
                    {
                        if (field.Contains("#      nhw"))
                        {
                            nhwIndex = idxCounter;
                        }
                        if (field.Contains("#     nchi"))
                        {
                            nchiIndex = idxCounter;
                        }
                        if (field.Contains("#     nram  iramform"))
                        {
                            nramIndex = idxCounter;
                        }
                        if (field.Contains("*Output Parameter"))
                        {
                            outputParamIndex = idxCounter;
                        }
                        if (field.Contains("# ipaccfou  ipacchis  ipaccspc  ipaccprb  ipaccmax ipaccoutd"))
                        {
                            ipaccfouIndex = idxCounter;
                        }
                        if (field.Contains("# nhaccout"))
                        {
                            nhaccoutIndex = idxCounter;
                        }
                        if (field.Contains("#  iprsfou   iprshis   iprsspc   iprsprb   iprsmax  iprsoutd"))
                        {
                            iprsfouIndex = idxCounter;
                        }
                        if (field.Contains("# idif2out  irad2out  idif3out  irad3out"))
                        {
                            idif2outIndex = idxCounter;
                        }
                        idxCounter++;
                    }
                }

                parser2.TextFieldType = FieldType.Delimited;
                parser2.SetDelimiters(",");
                int counter = 0;
                double nhw = 0;
                double nfn = 0;
                while (!parser2.EndOfData)
                {
                    string[] fields = parser2.ReadFields();

                    foreach (string field in fields)
                    {
                        string[] ssize = field.Split(null);
                        if (counter == nchiIndex + 1)
                        {
                            // Assign first value of this row to nchi
                            nchi = int.Parse(ssize[0]);
                        }
                        if (counter == nramIndex + 1)
                        {
                            // Assign values of this row to nram, iramform
                            nram = double.Parse(ssize[0]);
                            iramform = int.Parse(ssize[ssize.Length - 1]);
                        }
                        if (counter == nhwIndex + 1)
                        {
                            // Assign first value of this row to nhw
                            nhw = double.Parse(ssize[0]);
                        }
                        if (counter == 13)
                        {
                            // Assign first value of this row to nfn
                            nfn = double.Parse(ssize[0]);
                        }
                        if (counter > nhwIndex + 2 && counter < nchiIndex)
                        {
                            // Assign dictionary of vhw of this row to waveH
                            waveH.Add(new Dictionary<string, double>() { { "vhw", double.Parse(ssize[0]) } });
                        }
                        if (counter > nchiIndex + 2 && counter < nramIndex)
                        {
                            // Assign dictionary of vchi of this row to waveDir
                            waveDir.Add(new Dictionary<string, double>() { { "vchi", double.Parse(ssize[0]) } });
                        }
                        if (counter > nramIndex + 2 && counter < outputParamIndex)
                        {
                            // Assign dictionary of vram of this row to waveL
                            waveL.Add(new Dictionary<string, double>() { { "vram", double.Parse(ssize[0]) } });
                        }

                        counter++;
                    }
                }

                speed = nhw * nfn;
                /*for (int i = 0; i < waveDir.Count; i++)
                {
                    foreach (var kvp in waveDir[i])
                    {
                        Console.WriteLine(kvp.Key);
                        Console.WriteLine(kvp.Value);
                    }
                }*/

                foreach (var kvp in waveH[waveH.Count - 1])
                {
                    // Assign last waveH element's value to vhw
                    vhw = kvp.Value;
                }
            }

            // Read and Write xlsm
            using (var excelFileSource = new ExcelPackage(fileSource))
            using (var excelFileDestination = new ExcelPackage(fileDestination))
            {
                // Extract data from data sheet
                var dataSource = excelFileSource.Workbook.Worksheets[0];
                var responseData = new List<Dictionary<string, object>>();

                var dataCount = 2;
                while (true)
                {
                    var res = dataSource.Cells[4, dataCount].Value;
                    if (res == null)
                    {
                        break;
                    }
                    string[] arr = ((string)res).Split("_");

                    responseData.Add(new Dictionary<string, object>() { { "name", String.Join("_", arr[0..(arr.Length - 1)]) },
                    { "id", dataSource.Cells[5, dataCount].Value }
                    });

                    dataCount += 2;
                }

                // Write data to 計算設定 sheet
                var worksheetSource = excelFileSource.Workbook.Worksheets[1];
                worksheetSource.Cells[4, 5].Value = vhw;
                worksheetSource.Cells[5, 5].Value = alpp;
                worksheetSource.Cells[5, 9].Value = 1;
                worksheetSource.Cells[8, 3].Value = speed;


                string[] iramformList = new string[] { "λ/L", "√(L/λ)", "ω [rad/s]", "ω e[rad/s]", "T [s]", "T e[s]" };
                worksheetSource.Cells[11, 5].Value = iramformList[iramform];

                worksheetSource.Cells[14, 3].Value = nchi == 7 ? "Yes" : "No";
                worksheetSource.Cells[14, 5].Value = nram;

                for (var i = 0; i < waveDir.Count; i++)
                {
                    worksheetSource.Cells[17 + i, 3].Value = waveDir[i]["vchi"];
                }

                for (var i = 0; i < waveL.Count; i++)
                {
                    worksheetSource.Cells[17 + i, 5].Value = waveL[i]["vram"];
                }

                for (var i = 0; i < responseData.Count; i++)
                {
                    worksheetSource.Cells[12 + i, 9].Value = responseData[i]["name"];
                    worksheetSource.Cells[12 + i, 10].Value = responseData[i]["id"];
                    worksheetSource.Cells[12 + i, 11].Value = 1;
                }

                worksheetSource.Cells[8, 9].Value = responseData[0]["name"];

                excelFileSource.SaveAs(fileDestination);
            }
        }
    }
}
