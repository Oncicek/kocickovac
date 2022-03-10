using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace Kocickovac
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Čauky mňauky! Já jsem paní Kočička.");
            Console.WriteLine("Hledám klubíčko, tak chvilku počkej...");

            string localPath = Environment.CurrentDirectory;
            string[] sourceFile = Directory.GetFiles(localPath, "*.xlsx", SearchOption.TopDirectoryOnly);
            string directoryFinal = Directory.CreateDirectory(Path.Combine(localPath, "Result")).FullName;

            if (sourceFile.Length > 0)
            {
                Application excel = new Application()
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                //Source excel
                Workbook wb = excel.Workbooks.Open(sourceFile[0]);
                Worksheet sourceFirstSheet = wb.Sheets[1];

                //Get header
                Range header = sourceFirstSheet.get_Range("A1", "A1").EntireRow;

                int lastUsedRow = sourceFirstSheet.Cells.Find("*", System.Reflection.Missing.Value,
                                   System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                   XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
                                   false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                Console.WriteLine(String.Format("Našla jsem klubíčko se jménem {0}, které má {1} řádků. Po kolika řádcích ho mám rozmotat?", wb.Name, lastUsedRow - 1));

                string input = Console.ReadLine();
                int offset;

                while (!int.TryParse(input, out offset))
                {
                    Console.WriteLine("To není číslo, pitomče!");

                    input = Console.ReadLine();
                }

                int numOfFinalFiles = (int)Math.Ceiling((lastUsedRow - 1) / (double)offset);
                int firstRowDefault = 2;

                List<int> starts = new List<int>();
                List<int> ends = new List<int>();

                for (int j = 0; j < numOfFinalFiles; j++)
                {
                    starts.Add(firstRowDefault + (offset * j));

                    if (j < numOfFinalFiles - 1)
                    {
                        ends.Add((firstRowDefault - 1) + (offset * (j + 1)));
                    }
                    else
                    {
                        ends.Add(lastUsedRow);
                    }
                }

                for (int i = 0; i < numOfFinalFiles; i++)
                {
                    //New excel
                    Workbook newWb = excel.Workbooks.Add();
                    Worksheet firstSheet = newWb.Sheets[1];

                    //paste header
                    firstSheet.get_Range("A1", "A1").EntireRow.Value = header.Value;

                    Console.WriteLine(String.Format("Shazuji {0}. soubor z poličky...", i + 1));

                    Range copyRange = sourceFirstSheet.Rows[starts[i] + ":" + ends[i]];
                    copyRange.Copy(firstSheet.Rows[firstRowDefault + ":" + (offset + firstRowDefault)]);

                    string finalPathWithName = Path.Combine(directoryFinal, String.Format("Soubor_{0}", i + 1));

                    Console.WriteLine("Soubor spadnul do složky 'Result', tak si ho tam najdi.");

                    try
                    {
                        newWb.SaveAs(finalPathWithName);
                    }
                    catch (Exception)
                    {
                        newWb.Close();
                    }
                }

                Console.WriteLine("Hopky hopky domky.");

                try
                {
                    wb.Close();
                    excel.Quit();
                    CleanUp.KillExcel();
                }
                catch (Exception)
                {
                    CleanUp.KillExcel();
                }
            }
            else
            {
                Console.WriteLine("Nic tu není, darmožroute!");
            }
        }
    }
}
