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
            var fileSource = new FileInfo($"{pathDirectory}/../input/DLP-Prediction.xlsm");
            var fileDestination = new FileInfo($"{pathDirectory}/../output/DLP-Prediction1.xlsm");

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
            using (TextFieldParser parser = new TextFieldParser($"{pathDirectory}/../input/SHIPINP.DAT"))
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
            using (TextFieldParser parser = new TextFieldParser($"{pathDirectory}/../input/INPPARA.DAT"))
            using (TextFieldParser parser2 = new TextFieldParser($"{pathDirectory}/../input/INPPARA.DAT"))
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
                var responseData = new List<Dictionary<string, string>>();

                var dataCount = 2;
                while (true)
                {
                    var res = dataSource.Cells[4, dataCount].Value;
                    if (res == null)
                    {
                        break;
                    }
                    string[] arr = ((string)res).Split("_");

                    responseData.Add(new Dictionary<string, string>() { { "name", String.Join("_", arr[0..(arr.Length - 1)]) },
                    { "id", dataSource.Cells[5, dataCount].Value.ToString() }
                    });

                    dataCount += 2;
                }

                // Write data to 計算設定 sheet
                var worksheetSource = excelFileSource.Workbook.Worksheets[1];
                worksheetSource.Cells[4, 5].Value = vhw;
                worksheetSource.Cells[5, 5].Value = alpp;
                worksheetSource.Cells[5, 9].Value = 1;
                worksheetSource.Cells[8, 3].Value = speed;


                string[] iramformList = new string[] { "𝜆/L", "√(L/𝜆)", "ⲱ[rad/s]", "ⲱe[rad/s]", "T[s]", "Te[s]" };
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
                    worksheetSource.Cells[12 + i, 10].Value = int.Parse(responseData[i]["id"]);
                    worksheetSource.Cells[12 + i, 11].Value = 1;
                }

                worksheetSource.Cells[8, 9].Value = responseData[0]["name"];

                excelFileSource.SaveAs(fileDestination);
            }
        }
    }
}
