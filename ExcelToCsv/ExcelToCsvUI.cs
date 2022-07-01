using Microsoft.Office.Interop.Excel;
using System;
using System.IO;

namespace ExcelToCsv
{
    public class ExcelToCsvUI
    {
        public void Run()
        {
            ConvertExcel();
        }

        public void ConvertExcel()
        {
            bool keepRunning = true;

            while (keepRunning)
            {
                //Prompt User code

                Console.WriteLine(@"Enter the path to your .xlsx file (for example: C:\MyDocuments\MySpreadsheet.xlsx)");
                string fileLocation = Console.ReadLine();
                string directoryName = Path.GetDirectoryName(fileLocation);
                Console.WriteLine("Enter the number of rows in your spreadsheet (including the header row)");

                int numberOfRows = int.Parse(Console.ReadLine());

                int errorRow = 1;

                string fileNameWithoutExtension = Path.ChangeExtension(fileLocation, null);

                //Stream writer to write to txt document

                StreamWriter sw = new StreamWriter(fileNameWithoutExtension +".csv.txt");

                //Error speadsheet
                
                
                if (fileLocation == "" || numberOfRows == 0)
                {
                    
                    //Console.Clear();
                }
                else if(numberOfRows > 10000)
                {
                    //Under10kConvert(fileLocation, numberOfRows, directoryName, sw, errorRow);
                    Over10kConvert(fileLocation, numberOfRows, directoryName, sw, errorRow);
                    ConfirmationMessage();
                    keepRunning = false;

                }
                else
                {
                    Under10kConvert(fileLocation, numberOfRows, directoryName, sw, errorRow);
                    ConfirmationMessage();
                    keepRunning = false;


                }
            }
        }


        public void Under10kConvert(string fileLocation, int numberOfRows, string directoryName, StreamWriter sw, int errorRow)
        {

            //Open correct worksheets and create an error spreadsheet

            Application app = new Application();

            Workbook wb = app.Workbooks.Open($"{fileLocation}");

            Workbook errorWb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            


            Worksheet ws = wb.Worksheets[1];
            Worksheet errorWs = errorWb.Worksheets[1];


            int numberOfRowsUnder10k = numberOfRows;

            if(numberOfRows > 10001)
            {
                numberOfRowsUnder10k = 10001;
            }

            

            //Column Header Values

            string headerValues = "PID^Product id^Mfr Name^Mfr P/N^Price^COO^Short Description^UPC^UOM \n";

            sw.WriteLine(headerValues);

           

            //Get values of cells and assign to object

            for (double i = 2; i < numberOfRowsUnder10k; i++)
            {
                

                Console.WriteLine($"{i / numberOfRows * 100}% Completed");

                var rowOfData = ws.Rows[$"{i}"];

                Excel excel = new Excel();
                Csv csv = new Csv();

                //Current row in error.xlsx

                

                int notNumber = 0;
                //Assign cell data to an object
                
                bool tryParse = Int32.TryParse(rowOfData.Cells[1].Value.ToString(), out notNumber);

                if (tryParse)
                {
                    excel.PID = Convert.ToInt32(rowOfData.Cells[1].Value);
                    excel.ProductId = rowOfData.Cells[2].Value;
                    excel.MfrName = rowOfData.Cells[3].Value;
                    excel.MfrPN = rowOfData.Cells[4].Value;
                    excel.Cost = rowOfData.Cells[5].Value;
                    if (rowOfData.Cells[6].Value == "")
                    {
                        excel.COO = "TW";
                    }
                    else { excel.COO = rowOfData.Cells[6].Value; }
                    excel.ShortDescription = rowOfData.Cells[7].Value;
                    excel.UPC = rowOfData.Cells[8].Value;
                    if (rowOfData.Cells[9].Value == "")
                    {
                        excel.UOM = "EA";
                    }
                    else { excel.UOM = rowOfData.Cells[9].Value; };

                    //Convert excel data to csv data

                    csv.PID = excel.PID;
                    csv.ProductId = excel.ProductId;
                    csv.MfrName = excel.MfrName;
                    csv.MfrPN = excel.MfrPN;
                    csv.Price = excel.Cost * 1.2;
                    csv.COO = excel.COO;
                    csv.ShortDescription = excel.ShortDescription;
                    csv.UPC = excel.UPC;
                    csv.UOM = excel.UOM;

                    //CSV object to an array

                    string[] arrForCsv = { $"{csv.PID}", csv.ProductId, csv.MfrName, csv.MfrPN, $"{csv.Price}", csv.COO, csv.UPC, csv.UOM };


                    //write data to cell range
                    sw.WriteLine($"{arrForCsv[0]}^{arrForCsv[1]}^{arrForCsv[2]}^{arrForCsv[3]}^{arrForCsv[4]}^{arrForCsv[5]}^{arrForCsv[6]}^{arrForCsv[7]}\n");

                    
                    Console.Clear();
                    
                }

                else
                {
                    
                    

                    string[] errorSpeadsheetData = { Convert.ToString(rowOfData.Cells[1].Value) , Convert.ToString(rowOfData.Cells[2].Value), Convert.ToString(rowOfData.Cells[3].Value), Convert.ToString(rowOfData.Cells[4].Value), Convert.ToString(rowOfData.Cells[5].Value), Convert.ToString(rowOfData.Cells[6].Value), Convert.ToString(rowOfData.Cells[7].Value), Convert.ToString(rowOfData.Cells[8].Value), Convert.ToString(rowOfData.Cells[9].Value) };

                    errorWs.Rows[$"{errorRow}"] = errorSpeadsheetData;

                    errorRow++;
                   
                }



            }


            //Save and close all files
            errorWb.SaveAs(directoryName + @"\Error.xlsx");
            wb.Close();
            errorWb.Close();
            sw.Close();
            

            
        }

        public void ConfirmationMessage()
        {
            Console.WriteLine("Your file has been converted. Press any key to exit.");

            Console.ReadKey();
            
        }

        public void Over10kConvert(string fileLocation, int numberOfRows, string directoryName, StreamWriter sw, int errorRow)
        {
            //Open correct worksheets and create an error spreadsheet

            Application app = new Application();

            Workbook wb = app.Workbooks.Open($"{fileLocation}");

            Workbook errorWb = app.Workbooks.Open(directoryName+ @"\Error.xlsx");

            StreamWriter over10kSw = new StreamWriter(Path.ChangeExtension(fileLocation, null) + "(2).csv.txt");

            Worksheet ws = wb.Worksheets[1];
            Worksheet errorWs = errorWb.Worksheets[1];






            //Column Header Values

            string headerValues = "PID^Product id^Mfr Name^Mfr P/N^Price^COO^Short Description^UPC^UOM \n";

            over10kSw.WriteLine(headerValues);


            


            //Get values of cells and assign to object

            for (double i = 10001; i < numberOfRows; i++)
            {


                Console.WriteLine($"{i / numberOfRows * 100}% Completed");

                var rowOfData = ws.Rows[$"{i}"];

                Excel excel = new Excel();
                Csv csv = new Csv();

                

                int notNumber = 0;
                //Assign cell data to an object

                bool tryParse = Int32.TryParse(rowOfData.Cells[1].Value.ToString(), out notNumber);

                if (tryParse)
                {
                    excel.PID = Convert.ToInt32(rowOfData.Cells[1].Value);
                    excel.ProductId = rowOfData.Cells[2].Value;
                    excel.MfrName = rowOfData.Cells[3].Value;
                    excel.MfrPN = rowOfData.Cells[4].Value;
                    excel.Cost = rowOfData.Cells[5].Value;
                    if (rowOfData.Cells[6].Value == "")
                    {
                        excel.COO = "TW";
                    }
                    else { excel.COO = rowOfData.Cells[6].Value; }
                    excel.ShortDescription = rowOfData.Cells[7].Value;
                    excel.UPC = rowOfData.Cells[8].Value;
                    if (rowOfData.Cells[9].Value == "")
                    {
                        excel.UOM = "EA";
                    }
                    else { excel.UOM = rowOfData.Cells[9].Value; };

                    //Convert excel data to csv data

                    csv.PID = excel.PID;
                    csv.ProductId = excel.ProductId;
                    csv.MfrName = excel.MfrName;
                    csv.MfrPN = excel.MfrPN;
                    csv.Price = excel.Cost * 1.2;
                    csv.COO = excel.COO;
                    csv.ShortDescription = excel.ShortDescription;
                    csv.UPC = excel.UPC;
                    csv.UOM = excel.UOM;

                    //CSV object to an array

                    string[] arrForCsv = { $"{csv.PID}", csv.ProductId, csv.MfrName, csv.MfrPN, $"{csv.Price}", csv.COO, csv.UPC, csv.UOM };


                    //write data to cell range
                    over10kSw.WriteLine($"{arrForCsv[0]}^{arrForCsv[1]}^{arrForCsv[2]}^{arrForCsv[3]}^{arrForCsv[4]}^{arrForCsv[5]}^{arrForCsv[6]}^{arrForCsv[7]}\n");


                    Console.Clear();

                }

                else
                {



                    string[] errorSpeadsheetData = { Convert.ToString(rowOfData.Cells[1].Value), Convert.ToString(rowOfData.Cells[2].Value), Convert.ToString(rowOfData.Cells[3].Value), Convert.ToString(rowOfData.Cells[4].Value), Convert.ToString(rowOfData.Cells[5].Value), Convert.ToString(rowOfData.Cells[6].Value), Convert.ToString(rowOfData.Cells[7].Value), Convert.ToString(rowOfData.Cells[8].Value), Convert.ToString(rowOfData.Cells[9].Value) };

                    errorWs.Rows[$"{errorRow}"] = errorSpeadsheetData;

                    errorRow++;

                }



            }


            //Save and close all files
            
            wb.Close();
            errorWb.Close();
            over10kSw.Close();
        }

    }

}
