using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace salesTest
{
    class Program
    {
        bool flag = true; // 测试用例是否正常
        string filePath { get; set; }
        string resultPath = null;
        string fileName = null;

        public void readExcel()
        {

            try
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlworkbook = xlApp.Workbooks.Open(filePath);
                Excel._Worksheet xlworksheet = xlworkbook.Sheets[1];
                Excel.Range xlRange = xlworksheet.UsedRange;

                int rowCnt = xlRange.Rows.Count;
                int colCnt = xlRange.Columns.Count;
                xlRange.Cells[1, 2] = DateTime.Now;
                xlRange.Cells[1, 5] = "彭嘉琦 李威";

                for (int i = 4; i <= rowCnt; i++)
                {
                    if (xlRange.Cells[i, 1] == null || xlRange.Cells[i, 1].Value2 == null)
                        break;
                    string result = null;
                    try
                    {
                        result = countProfit(int.Parse(xlRange.Cells[i, 2].Text), int.Parse(xlRange.Cells[i, 3].Text), int.Parse(xlRange.Cells[i, 4].Text));

                    }
                    catch (FormatException)
                    {
                        //Console.WriteLine(fe.Message);
                        flag = false;
                        result = "输入类型不正确";
                    }
                    if (!flag) //用例输入异常
                    {
                        xlRange.Cells[i, 7] = "error";
                        //xlRange.Cells[i, 9] = result;
                        xlRange.Cells[i, 8] = (string.Compare(xlRange.Cells[i, 6].Text, "error") == 0) ? "T" : "F";
                    }
                    else
                    {
                        xlRange.Cells[i, 7] = result;
                        xlRange.Cells[i, 8] = (string.Compare(xlRange.Cells[i, 6].Text, result) == 0) ? "T" : "F";
                    }

                }

                fileName = System.IO.Path.GetFileNameWithoutExtension(filePath) + "结果" + DateTime.Now.ToString("yyyyMMddHHmmss");
                resultPath = System.IO.Path.GetDirectoryName(filePath) +"\\";
                xlworkbook.SaveAs(resultPath + fileName + ".xlsx", Type.Missing, "", "", Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, 1, false, Type.Missing, Type.Missing, Type.Missing);
                xlworkbook.Close();
                xlApp.Quit();
                Console.WriteLine(resultPath + fileName + ".xlsx");
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                //Console.WriteLine(ex.Message);
            }

        }

        private string countProfit(int peripheral, int mainframe, int display)
        {
            flag = true;
            double profit = 0;
            if (peripheral < 0 || mainframe < 0 || display < 0)
            {
                flag = false;
                return "销售量为负数";
            }
            if (peripheral < 1 || mainframe < 1 || display < 1)
            {
                flag = false;
                return "销售机器少于一台";
            }
            if (peripheral > 90 || mainframe > 70 || display > 80)
            {
                flag = false;
                return "超过最大销售额";
            }
            double profitTmp = peripheral * 25 + mainframe * 45 + display * 30;
            if (profitTmp < 1000)
            {
                profit = profitTmp * 0.05;
            }
            else if (profitTmp >= 1000 && profitTmp < 1800)
            {
                profit = profitTmp * 0.1;
            }
            else
            {
                profit = profitTmp * 0.15;
            }
            return profit.ToString();
        }

        static void Main(string[] args)
        {
            Program p = new Program();
            if (args.Length != 0)
            {
                p.filePath = args[0].Trim();
                p.readExcel();
            }
            else
            {
                p.filePath = @"C:\Users\CCMEOW\Desktop\sales边界值.xlsx";
                p.readExcel();
            }


        }
    }
}
