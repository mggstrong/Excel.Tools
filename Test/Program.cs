using Excel.Tools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory + "Templates/temp1.xls";
            var abcExcel = AppDomain.CurrentDomain.BaseDirectory + "Templates/abc.xls";
            var w = new ExcelWriter();
            //try
            //{


            //    Console.WriteLine("---read byte from template---path:" + baseDir);
            //    var result = w.Excute(baseDir, new[] { "v1", "v2", "v3", "v4", "v5" });
            //    Console.WriteLine("---write byte to file---");
            //    File.WriteAllBytes(, result);
            //    Console.WriteLine("---success---");
            //}
            //catch (Exception ex)
            //{

            //    Console.WriteLine("---error:" + ex.Message + "---");
            //}

            var r = new ExcelReader();
            var m = r.Excute(abcExcel);
            for (int i = 0; i < m.GetLength(0); i++)
            {
                string rowValue = "";
                for (int j = 0; j < m.GetLength(1); j++)
                {
                    rowValue += (m[i, j] + ",");
                }
                Console.WriteLine(rowValue);
            }
            Console.ReadLine();
        }
    }
}
