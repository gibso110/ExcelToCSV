using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

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

                if(fileLocation == "")
                {
                    Console.Clear();
                }
                else
                {
                    Application app = new Application();
                    Workbook wb = app.Workbooks.Open($"{fileLocation}");

                    Worksheet ws = wb.Worksheets[1];

                    string fileNameWithoutExtension = System.IO.Path.ChangeExtension(fileLocation, null);

                    //Column Header Values

                    string[] headerValues = new string[] { "PID", "Product id", "Mfr Name", "Mfr P/N", "Price", "COO", "Short Description", "UPC", "UOM" };

                    ws.get_Range("A1", "I1").Value = headerValues;


                    //Get values of cells and assign to object

                    for(int i = 2; i < 10000; i++)
                    {
                       var rowOfData = ws.Rows[$"{i}"];

                        Excel excel = new Excel();
                        Csv csv = new Csv();

                        //Assign cell data to an object

                        excel.PID = Convert.ToInt32(rowOfData.Cells[1].Value);
                        excel.ProductId = rowOfData.Cells[2].Value;
                        excel.MfrName = rowOfData.Cells[3].Value;
                        excel.MfrPN = rowOfData.Cells[4].Value;
                        excel.Cost = rowOfData.Cells[5].Value;
                        if(rowOfData.Cells[6].Value == "")
                        {
                            excel.COO = "TW";
                        }
                        else { excel.COO = rowOfData.Cells[6].Value; }
                        excel.ShortDescription = rowOfData.Cells[7].Value;
                        excel.UPC = rowOfData.Cells[8].Value;
                       if(rowOfData.Cells[9].Value == "")
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

                        string[] arrForCsv = { $"{csv.PID}", csv.ProductId, csv.MfrName, csv.MfrPN , $"{csv.Price}", csv.COO , csv.UPC, csv.UOM };


                        //write data to cell range
                        rowOfData.Value = arrForCsv;




                    }


                    //Save and Convert to CSV

                    wb.SaveAs($"{fileNameWithoutExtension}", XlFileFormat.xlCSV, Type.Missing,Type.Missing, false);

                    Console.WriteLine("Your file has been converted. Press any key to exit.");

                    Console.ReadKey();
                    keepRunning = false;
                }
            }
        }


    }
}
