using System;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;

namespace PDFPageCounter2000
{
    
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Please Enter your Username...");
            string username = Console.ReadLine().ToString();
            Console.WriteLine("Please Enter your Password...");
            string password = Console.ReadLine().ToString();
            //D:\!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Forbatches1through5\UsernamesPasswordsForProgramTesting.txt
            string[] lines = File.ReadAllLines(@"D:\!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Forbatches1through5\UsernamesPasswordsForProgramTesting.txt").Skip(1).ToArray();
            bool result = false;
            foreach(var line in lines)
            {
                var firstValue = line.Split(',');
                string usernameInfo = firstValue[0];
                string passwordInfo = firstValue[1];

                if (usernameInfo == username)
                {
                    if (passwordInfo == password)
                    {
                        result = true;
                        break;
                    }
                    else
                    {
                        result = false;
                        continue;
                    }
                }
                else
                {
                    result = false;
                    continue;
                }

            }
            

            if (result == true)
            {
                Console.WriteLine("Access Granted...");
                Excel.Application xlApplication = new Excel.Application();

                Console.WriteLine("Clocking In(1) or Clocking Out(2)?");
                string clockInOrOut = Console.ReadLine();
                int columnDenotation = 0;
                int row = 0;

                Excel.Workbook xlWorkbook = xlApplication.Workbooks.Open("D:\\!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Forbatches1through5\\TimeSheetProgramTest_10.xlsx");
                //Console.WriteLine(username);
                Excel.Worksheet xlWorksheet = new Excel.Worksheet();
                bool result2 = false;
                foreach(Excel.Worksheet sheet in xlWorkbook.Sheets)
                {
                    if(sheet.Name == username)
                    {
                        xlWorksheet = sheet;
                        result2 = true;
                        break;
                    }
                    else
                    {
                        result2 = false;
                    }
                }

                if(result2 != true)
                {
                    xlWorksheet = xlWorkbook.Sheets.Add();

                    xlWorksheet.Name = username.ToString();
                    
                }
                xlWorkbook.Save();


                if (xlWorksheet.Cells[1, 1].Value == null)
                {
                    xlWorksheet.Cells[1, 1].Value = "Date";
                }
                if (xlWorksheet.Cells[1, 2].Value == null)
                {
                    xlWorksheet.Cells[1, 2].Value = "Time-In";
                }
                if (xlWorksheet.Cells[1, 3].Value == null)
                {
                    xlWorksheet.Cells[1, 3].Value = "Time-Out";
                }
                if (xlWorksheet.Cells[1, 4].Value == null)
                {
                    xlWorksheet.Cells[1, 4].Value = "Elapsed";
                }

                if (clockInOrOut == 1.ToString())
                {
                    columnDenotation = 2;
                }
                if (clockInOrOut == 2.ToString())
                {
                    columnDenotation = 3;
                }


                //Date Injector
                for (int j = 2; j <= 100000; j++)
                {
                    if (xlWorksheet.Cells[j, 1].Value == null)
                    {
                        xlWorksheet.Cells[j, 1].Value = DateTime.Now.ToShortDateString().ToString();
                        xlWorkbook.Save();
                        if (xlWorksheet.Cells[j, columnDenotation].Value == null)
                        {
                            xlWorksheet.Cells[j, columnDenotation].Value = DateTime.Now.ToLongTimeString();
                            row = j;
                            xlWorkbook.Save();
                            
                            break;
                        }
                    }
                    else
                    {
                        if (xlWorksheet.Cells[j, columnDenotation].Value == null)
                        {
                            xlWorksheet.Cells[j, columnDenotation].Value = DateTime.Now.ToLongTimeString();
                            row = j;
                            xlWorkbook.Save();
                            
                            break;
                        }
                    }
                }

                

                xlWorkbook.Save();
                xlWorkbook.Close();


                Console.WriteLine("Press Any Key to ESC...");
                Console.ReadKey();

            }
            else
            {
                Console.WriteLine("Access Denied, Invalid Credentials...");
                Console.WriteLine("Press Any Key to ESC...");
                Console.ReadKey();
            }

        }
    }
}
