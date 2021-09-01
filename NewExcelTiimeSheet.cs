using System;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace PDFPageCounter2000
{
    
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Please Enter your Username...(If New User, Type 'New User'");
            string username = Console.ReadLine().ToString();
            string password = null;
            if (username != "New User")
            {
                Console.WriteLine("Please Enter your Password...");
                password = Console.ReadLine().ToString();
            }

            //D:\!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Forbatches1through5\UsernamesPasswordsForProgramTesting.txt
            string[] lines = File.ReadAllLines(@"D:\!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Forbatches1through5\UsernamesPasswordsForProgramTesting.txt").Skip(1).ToArray();
            bool result = false;

            ////Credential checking begins here
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
                if (username == "New User")
                {
                    Console.WriteLine("Please enter your descriptive username(FirstnameLastname");
                    username = Console.ReadLine();
                    Console.WriteLine("Please enter a secure password");
                    password = Console.ReadLine();
                    using (StreamWriter swr = new StreamWriter(@"D:\!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Forbatches1through5\UsernamesPasswordsForProgramTesting.txt", true))
                    {
                        swr.Write(Environment.NewLine);
                        swr.Write(String.Format((username + "," + password)));
                    }
                    Console.WriteLine("Success! Please continue with application...");
                    result = true;
                    break;
                }
                else
                {
                    result = false;
                    continue;
                }
            }
            ////Credential checking ends here, sorta

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

                ////Logic for setting up unique sheets for each user starts here
                foreach(Excel.Worksheet sheet in xlWorkbook.Sheets)
                {
                    if(sheet.Name == username)
                    {
                        xlWorksheet = sheet;
                        xlWorkbook.Save();
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
                    xlWorkbook.Save();
                    xlWorksheet.Name = username.ToString();
                    xlWorkbook.Save();
                }
                xlWorkbook.Save();
                ////Logic for setting up unique sheets for each user ends here

                ////Start Column headers for user's sheet
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
                ////End of Setting column headers for user's sheet

                ////Start of continuing the clockin or clockout setting
                if (clockInOrOut == 1.ToString())
                {
                    columnDenotation = 2;
                }
                if (clockInOrOut == 2.ToString())
                {
                    columnDenotation = 3;
                }
                ////End of continuing the clockin or clockout setting

                ////Pasting values into correct columns and rows begins here
                for (int j = 2; j <= 100000; j++)
                {
                    if (xlWorksheet.Cells[j, 1].Value == null) //Date Injector
                    {
                        xlWorksheet.Cells[j, 1].Value = DateTime.Now.ToShortDateString().ToString();
                        xlWorkbook.Save();
                        if (xlWorksheet.Cells[j, columnDenotation].Value == null) //Time Injector
                        {
                            xlWorksheet.Cells[j, columnDenotation].Value = DateTime.Now.ToLongTimeString();
                            row = j;
                            xlWorkbook.Save();
                            
                            break;
                        }
                    }
                    else
                    {
                        if (xlWorksheet.Cells[j, columnDenotation].Value == null) //Time Injector
                        {
                            xlWorksheet.Cells[j, columnDenotation].Value = DateTime.Now.ToLongTimeString();
                            row = j;
                            xlWorkbook.Save();
                            
                            break;
                        }
                    }
                }
                ////Pasting values into correct columns and rows ends here



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
