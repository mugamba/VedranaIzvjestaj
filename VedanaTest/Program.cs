using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Packaging;
using System.Numerics;

namespace VedranaIlirija
{
    class Program
    {
        static void Main(string[] args)
        {
            //var validList = new List<String>() {
            //    };

            BigInteger bigIntFromDouble = new BigInteger(179032.6541);

            var validList =  File.ReadAllLines("InputKonta.txt").ToList();

            var files = Directory.GetFiles(@"C:\Temp\VedranaIlirija", "*.xls");

            foreach (var file in files)
            {

                var name = Path.GetFileNameWithoutExtension(file);
                var dir = Path.GetDirectoryName(file);
                Application app = new Application();
                Workbook wb = app.Workbooks.Open(file);
                wb.SaveAs(Path.Combine(dir, name + ".txt"), XlFileFormat.xlCurrentPlatformText);
                wb.Close(false);
                app.Quit();

            }


            files = Directory.GetFiles(@"C:\Temp\VedranaIlirija", "*.txt");

            foreach (var file in files)
            {
                var enc = Encoding.GetEncoding("Windows-1250");

                var name = Path.GetFileNameWithoutExtension(file);
                var dir = Path.GetDirectoryName(file) + @"\Filtered";

                Directory.CreateDirectory(dir);

                var lines = File.ReadAllLines(file, enc);

                var builder = new StringBuilder();
                string sep = @"\t";

                foreach (var line in lines)
                {
                    var splitedline = line.Split('\t');

                    splitedline = splitedline.Where(o => !String.IsNullOrEmpty(o)).Select(o => o.Trim('"')).ToArray();
                    var test = splitedline[0];


                    //splitedline[0] = splitedline[0];
                    //splitedline[1] = splitedline[1];



                    if (validList.Contains(test))
                        builder.AppendLine(String.Join("\t", splitedline));

                }




                File.WriteAllText(Path.Combine(dir, name + ".txt"), builder.ToString(), enc);

            }

                
            
            files = Directory.GetFiles(@"C:\Temp\VedranaIlirija\Filtered", "*.txt");


            foreach (var file in files)
            {

                var name = Path.GetFileNameWithoutExtension(file);
                var dir = Path.GetDirectoryName(file);
                Application app = new Application();
                Workbook wb = app.Workbooks.Open(file);
                wb.SaveAs(Path.Combine(dir, name + ".xls"), XlFileFormat.xlExcel7);
                wb.Close(false);
                app.Quit();

            }




        }
    }
}
