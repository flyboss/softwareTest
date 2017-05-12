using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;
namespace 话费
{
    class Program
    {
        //折扣率
        private static decimal myDiscount(int tellTime,int count)
        {
            decimal myDiscount=0;
            if(tellTime<=60)
            {
                if(count <= 1)
                    myDiscount = 0.01M;
            }
            else if(tellTime<=120)
            {
                if (count <= 2)
                    myDiscount = 0.015M;
            }
            else if(tellTime<=180)
            {
                if (count <= 3)
                    myDiscount = 0.02M;
            }
            else if(tellTime<=300)
            {
                if (count <= 3)
                    myDiscount = 0.025M;
            }
            else
            {
                if (count <= 6)
                    myDiscount = 0.03M;
            }
            return myDiscount;
        }
        private static decimal mySpend(int tellTime,int count)
        {
            decimal discount = myDiscount(tellTime, count);
            decimal sum = 25M + tellTime * 0.15M * (1M-discount);
            return sum;
        }
        //判断输入数据是否正确
        private static decimal myInit(int[] testData)
        {
            int tellTime = testData[0];
            int count = testData[1];

            if(tellTime<=0||tellTime> 44640|| count<0||count>11)
            {
                return -1M;
            }
            return mySpend(tellTime, count);
        }
        /*
        public DataSet ExcelToDS(string Path)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties=Excel 8.0;";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataSet ds = null;
            strExcel = "select * from [sheet1$]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            ds = new DataSet();
            myCommand.Fill(ds, "table1");
            return ds;
        }*/
        private static void test(string path)
        {
            Application xlApp = new Application();
            xlApp.Visible = false;
            Workbook xlworkbook = xlApp.Workbooks.Open(path);
            Worksheet xlworksheet = xlworkbook.Sheets[1]; 
            Range xlRange = xlworksheet.UsedRange;

            int rowCnt = xlRange.Rows.Count;
            int colCnt = xlRange.Columns.Count;
            xlRange.Cells[1, 2] = DateTime.Now;
            xlRange.Cells[1, 5] = "彭嘉琦 李威";

            Range c1 = (Range)xlworksheet.Cells[4, 1];
            Range c2 = (Range)xlworksheet.Cells[rowCnt, colCnt];
            Range rng = xlworksheet.get_Range(c1, c2);
            object[,] exceldata = (object[,])rng.get_Value(XlRangeValueDataType.xlRangeValueDefault);
            for (int i = 1; i <= rowCnt - 3; i++)
            {
                int[] testData = { Convert.ToInt32(exceldata[i, 2]), Convert.ToInt32(exceldata[i, 3]) };
                string exStr = Convert.ToString(exceldata[i, 4]);
                decimal expect;
                if(exStr.Equals("error"))
                {
                    expect = -1M;
                }
                else
                {
                    expect = Convert.ToDecimal(exceldata[i, 4]);
                }
                decimal realout = myInit(testData);
                if(realout.Equals(-1M))
                {
                    xlRange.Cells[i + 3, 5] = "error";
                }
                else
                {
                    xlRange.Cells[i + 3, 5] = realout;
                }
                if (realout.Equals(expect))
                {
                    xlRange.Cells[i + 3, 6] = "T";
                }
                else
                {
                    xlRange.Cells[i + 3, 6] = "F";
                }
            }
            char sp = '\\';
            string[] road = path.Split(sp);
            string newPath = "";
            for (int i = 0; i < road.Length - 1; i++)
            {
                newPath += road[i] + "\\";
            }
            newPath += "话费测试报告" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            xlworkbook.SaveAs(newPath);
            xlworkbook.Close();
            xlApp.Quit();
            Console.WriteLine(newPath);
        } 
        static void Main(string[] args)
        {
            if (args.Length != 0)
              test(args[0].Trim());
        }
    }
}
