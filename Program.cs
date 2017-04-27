using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace nextday
{
    class Program
    {
        private static int[] leap = new int[] { 0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
        private static int[] noleap = new int[] { 0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
        private static bool myCheck(int[] testDate, List<string> expect, List<string> realout)
        {
            int year = testDate[0];
            int month = testDate[1];
            int day = testDate[2];
            if (year < 1 || year > 10000 || month < 1 || month > 12 || day < 1 || day > 31)
            {
                if (compareExpect(expect, realout, "error", "error", "error", "日期不符合规定"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            bool run = myrun(year);
            int[] monthP;
            if (run == true)
            {
                monthP = leap;
            }
            else
            {
                monthP = noleap;
            }
            if (monthP[month] < day)
            {
                if (compareExpect(expect, realout, "error", "error", "error", "日期超出了月份最大日期"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else if (monthP[month] == day)
            {
                day = 1;
                month++;
                if (month > 12)
                {
                    month = 1;
                    year++;
                }
            }
            else
            {
                day++;
            }
            if (compareExpect(expect, realout, year.ToString(), month.ToString(), day.ToString(), "日期错误"))
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        private static bool compareExpect(List<string> expect, List<string> realout, string year, string month, string day, string error)
        {
            realout.Add(year);
            realout.Add(month);
            realout.Add(day);
            if (expect[0] == realout[0] && expect[1] == realout[1] && expect[2] == realout[2])
            {
                realout.Add("T");
                return true;
            }
            else
            {
                realout.Add("F");
                realout.Add(error);
                return false;
            }
        }
        private static bool myrun(int year)
        {
            if (year % 4 == 0 && year % 100 != 0 || year % 400 == 0)
                return true;
            else
                return false;
        }
        private static void test(string path)
        {
            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = false;
            Excel.Workbook xlworkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlworksheet = xlworkbook.Sheets[1];
            Excel.Range xlRange = xlworksheet.UsedRange;

            int rowCnt = xlRange.Rows.Count;
            int colCnt = xlRange.Columns.Count;
            xlRange.Cells[1, 2] = DateTime.Now;
            xlRange.Cells[1, 5] = "彭嘉琦 李威";

            Excel.Range c1 = (Excel.Range)xlworksheet.Cells[4, 1];
            Excel.Range c2 = (Excel.Range)xlworksheet.Cells[rowCnt, colCnt];
            Excel.Range rng = (Microsoft.Office.Interop.Excel.Range)xlworksheet.get_Range(c1, c2);
            object[,] exceldata = (object[,])rng.get_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault);
            for (int i = 1; i <= rowCnt - 3; i++)
            {
                int[] testDate = { Convert.ToInt32(exceldata[i, 2]), Convert.ToInt32(exceldata[i, 3]), Convert.ToInt32(exceldata[i, 4]) };
                List<string> expect = new List<string>();
                expect.Add(Convert.ToString(exceldata[i, 5]));
                expect.Add(Convert.ToString(exceldata[i, 6]));
                expect.Add(Convert.ToString(exceldata[i, 7]));
                List<string> realout = new List<string>();
                myCheck(testDate, expect, realout);
                for (int j = 0; j < realout.Count(); j++)
                {
                    xlRange.Cells[i + 3, j + 8] = realout[j];
                }

            }
            char sp = '\\';
            string[] road = path.Split(sp);
            string newPath="";
            for (int i = 0; i < road.Length-1 ; i++)
            {
                newPath += road[i] + "\\";
            }
            newPath+= "nextday测试报告"+ DateTime.Now.ToString("yyyyMMddHHmmss")+".xlsx";
            xlworkbook.SaveAs(newPath);
            xlworkbook.Close();
            xlApp.Quit();
            Console.WriteLine(newPath);
        }
        static void Main(string[] args)
        {
            if(args.Length!=0)
            test(args[0].Trim());
        }
    }
}
