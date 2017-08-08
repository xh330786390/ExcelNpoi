using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelNpoi.Tests
{
    class Program
    {
        static void Main(string[] args)
        {

            DataTable dt = new DataTable();
            dt.Columns.Add("name1");
            dt.Columns.Add("name2");
            dt.Columns.Add("name3");
            for (int i = 1; i < 1000; i++)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < 3; j++)
                {
                    dr[j] = "row" + i.ToString() + ": columne" + j.ToString();
                }
                dt.Rows.Add(dr);
            }

            ExcelNpoi.IExcelNpoi exl2003 = new ExcelNpoi.ExcelNpoi2003();
            exl2003.ExportToExcel(dt, @"E:\2003.xls", "钢联数据", "ABCDEF", 0);
            Console.WriteLine("2003导出完成");

            ExcelNpoi.IExcelNpoi exl2007 = new ExcelNpoi.ExcelNpoi2007();
            exl2007.ExportToExcel(dt, @"E:\2007.xlsx", "钢联数据", "ABCDEF", 0);
            Console.WriteLine("2007导出完成");
            Console.Read();
        }
    }
}
