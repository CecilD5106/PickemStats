using Microsoft.Office.Interop.Excel;
using System;

namespace PickemStats
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = "E:\\Code\\VSCode\\Node\\CFB01\\2019CFPickem.xlsx";
            Application excel = new Application();
            Workbook wb = excel.Workbooks.Open(path);

            try
            {
                Worksheet wsCurPick = wb.Worksheets["CurPick"];
                Worksheet wsWeekPick = wb.Worksheets[wsCurPick.Cells[1, 1].Value];

                int i = 5;
                while (wsWeekPick.Cells[i, 1].Value != "X")
                {
                    if (wsWeekPick.Cells[i, 1].Value != "N")
                    {
                        Worksheet wsTeamStat = wb.Worksheets[wsWeekPick.Cells[i, 1].Value];
                        //Transfer statistics from Team sheet to Weekly pickem sheet
                        wsWeekPick.Cells[i, 2].Value = wsTeamStat.Cells[5, 1].Value;
                        wsWeekPick.Cells[i, 3].Value = wsTeamStat.Cells[5, 2].Value;
                        wsWeekPick.Cells[i, 6].Value = wsTeamStat.Cells[5, 4].Value;
                        wsWeekPick.Cells[i, 8].Value = wsTeamStat.Cells[5, 6].Value;
                        wsWeekPick.Cells[i, 10].Value = wsTeamStat.Cells[5, 8].Value;
                        wsWeekPick.Cells[i, 12].Value = wsTeamStat.Cells[5, 10].Value;
                        wsWeekPick.Cells[i, 15].Value = wsTeamStat.Cells[5, 12].Value;
                        wsWeekPick.Cells[i, 17].Value = wsTeamStat.Cells[5, 14].Value;
                        wsWeekPick.Cells[i, 19].Value = wsTeamStat.Cells[5, 16].Value;
                        //Find the last in the games stats section
                        int j = 12;
                        while (wsTeamStat.Cells[j, 9].Value != null)
                        {
                            j++;
                        }
                        j--;

                        int iWin3 = 0;
                        int iWin5 = 0;
                        int l = 0;
                        for (var k = 0; k < 5; k++ )
                        {
                            if (j > 11)
                            {
                                if (l < 3)
                                {
                                    iWin3 += wsTeamStat.Cells[j, 9].Value;
                                    iWin5 += wsTeamStat.Cells[j, 9].Value;
                                    l++;
                                    j--;
                                }
                                else
                                {
                                    iWin5 += wsTeamStat.Cells[j, 9].Value;
                                    j--;
                                }
                            }
                        }

                        wsWeekPick.Cells[i, 23].Value = iWin3;
                        wsWeekPick.Cells[i, 24].Value = iWin5;
                    }

                    i++;
                }

                wb.Save();
                excel.Quit();
            }
            catch (Exception e)
            {
                excel.Quit();
                Console.WriteLine(e.ToString());
                throw;
            }
            finally
            {
                excel.Quit();
            }
        }
    }
}
