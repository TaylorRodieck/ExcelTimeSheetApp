using System;
using System.Text;
using System.Collections.Generic;
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
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Please Enter your Username... (If New User, Type 'New User')");
            string username = Console.ReadLine().ToString();
            
            string password = null;
            if (username != "New User") ////Allows new users to bypass having to input a password that they have yet to make
            {
                Console.WriteLine("Please Enter your Password...");
                password = Console.ReadLine().ToString();
            }
            List<string> tempLines = new List<string>();
            

            ////Checks for credentials txt file existence and possible creation begins here
            if (File.Exists(@"D:\!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Forbatches1through5\ExcelAppTester_1\UsernamesPasswordsForProgramTesting.txt") == true)
            {
                tempLines = File.ReadAllLines(@"D:\!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Forbatches1through5\ExcelAppTester_1\UsernamesPasswordsForProgramTesting.txt").Skip(1).ToList<string>();
            }
            else
            {
                using (StreamWriter sw1 = File.CreateText(@"D:\!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Forbatches1through5\ExcelAppTester_1\UsernamesPasswordsForProgramTesting.txt"))
                {
                    sw1.WriteLine("Usernames,Passwords");
                    //sw1.Close();
                }
                tempLines = File.ReadAllLines(@"D:\!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Forbatches1through5\ExcelAppTester_1\UsernamesPasswordsForProgramTesting.txt").Skip(1).ToList<string>();
            }
            
            bool result = false;
            string[] lines = tempLines.ToArray();
            ////Checks for credentials txt file existence and possible creation ends here
            

            ////Credential checking/setting begins here
            if (username == "New User") ///Creates account for a new user, and sets result to true to allow them to move on...
            {
                Console.WriteLine("Please enter your descriptive username(FirstnameLastname");
                username = Console.ReadLine();
                username = Encrypt(username);

                Console.WriteLine("Please enter a secure password");
                password = Console.ReadLine();
                password = Encrypt(password);
                
                File.AppendAllText(@"D:\!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Forbatches1through5\ExcelAppTester_1\UsernamesPasswordsForProgramTesting.txt", (username + "," + password));
                Console.WriteLine("Success! Please continue with application...");
                result = true;
                
            }
            username = Encrypt(username); /// Encrypts username for returning users
            password = Encrypt(password); /// Encrypts password for returning users

            if(result == false) /// result should only be false for users with accounts at this point
            {
                foreach (var line in lines)
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
                            result = false; ///This line is for a wrong password but correct username
                            continue;
                        }
                    }
                    else ///This line is for a wrong username
                    {
                        result = false;
                        continue;
                    }
                }
            }
            
            ////Credential checking/setting ends here, sorta

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
                foreach (Excel.Worksheet sheet in xlWorkbook.Sheets)
                {
                    if (sheet.Name == username)
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

                if (result2 != true)
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
        public static string Encrypt(string userInput) ////Encryption involves a basic conversion to hexadecimal for all letters and insertion of *
        {                                              ////in between letters for added weirdness
            byte[] bA = Encoding.Default.GetBytes(userInput);
            var hexString = BitConverter.ToString(bA);
            hexString = hexString.Replace("-", "*");
            //Console.WriteLine(hexString.ToString());
            return hexString.ToString();
        }

        public static string Decrypt(string stringToDecrypt)
        {
            List<string> catchArray = new List<string>();
            string userInput = stringToDecrypt;
            userInput = userInput.Replace("*", "-");
            string[] userInputArray = userInput.Split('-');
            //Console.WriteLine("Decryption Result Below...");
            //Console.WriteLine();
            foreach (string val in userInputArray)
            {
                int value = Convert.ToInt32(val, 16);
                string stringValue = Char.ConvertFromUtf32(value);
                char charValue = (char)value;
                catchArray.Add(stringValue);
                //Console.Write(stringValue);
            }
            
            //Console.WriteLine();
            //Console.WriteLine();
            return catchArray.Aggregate((i,j) => i +j);
        }
    }
}
