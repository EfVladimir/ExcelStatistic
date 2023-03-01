using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ConsoleApp7
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string path = "Test.xlsx";
            using (var package = new ExcelPackage(path))
            {
                var sheet = package.Workbook.Worksheets[0];
                var rows = sheet.Dimension.Rows;
                int M = 0, 
                    W = 0, 
                    N = 0;
                for (int i = 2; i <= rows; i++)
                {
                    if (sheet.Cells[i, 1].Value != null)
                    {
                        int temp = Convert.ToInt32(sheet.Cells[i, 1].Value.ToString().Substring(6, 1));
                        if (temp == 0)
                        {
                            N++;
                            sheet.Cells[i, 2].Value = "Н";
                        }
                        else if (temp % 2 == 0)
                        {
                            sheet.Cells[i, 2].Value = "Ж";
                            W++;
                        }
                        else
                        {
                            M++;
                            sheet.Cells[i, 2].Value = "М";
                        }
                    }
                    else 
                        break;
                }
                sheet.Cells[2,5].Value = "Женщины";
                sheet.Cells[2,6].Value = "Мужчины";
                sheet.Cells[2,7].Value = "Неопределено";
                sheet.Cells[3, 5].Value = W;
                sheet.Cells[3, 6].Value = M;
                sheet.Cells[3, 7].Value = N;
                package.Save();
            }
            Console.WriteLine("End");
            Console.ReadLine();
        }
    }
}
