using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.NetworkInformation;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using IExcelDataReader = ExcelDataReader.ExcelDataSetConfiguration;

namespace PingExcel_2._0
{
    class Program
    {
        static void Main(string[] args)
        {

            Run();
            Console.ReadKey();

        }

        private static async void Run()
        {
            string path;

            Console.Write("Введите путь, указав имя файла: ");
            path = Convert.ToString(Console.ReadLine());

            int i;
            int j = 1;
            int N;

            Console.Write("Количество строк: ");
            N = Convert.ToInt32((Console.ReadLine()));


            StreamReader f = new StreamReader(path);
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"" + path + "", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            for (j = 6; j <= 8; j++)
            {
                for (i = 3; i <= N; i++)
                {
                    string pinger = ObjWorkSheet.Cells[i, j].Text;
                    if (ObjWorkSheet.Rows[i].Text == null)
                    {
                        i++;
                    }

                    Ping pingSender = new Ping();

                    string data = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
                    byte[] buffer = Encoding.ASCII.GetBytes(data);
                    int timeout = 20000;

                    PingOptions options = new PingOptions(64, true);

                    try
                    {
                        PingReply reply = await pingSender.SendPingAsync(pinger);

                        if (reply.Status == IPStatus.Success)
                        {
                            Console.WriteLine(string.Format("{0} {1}", reply.Address, reply.Status, timeout));
                            (ObjWorkSheet.Cells[i, j] as Excel.Range).Interior.ColorIndex = 4;

                        }
                        else
                        {
                            Console.WriteLine(string.Format("{0}  {1}", reply.Address, reply.Status, timeout));
                            (ObjWorkSheet.Cells[i, j] as Excel.Range).Interior.ColorIndex = 3;
                        }
                    }
                    catch (Exception)
                    {
                        (ObjWorkSheet.Cells[i, j] as Excel.Range).Interior.ColorIndex = 6;
                    }
                }
            }
            Console.WriteLine("ЗАВЕРШЕНО");
            Console.ReadLine();
            ObjWorkExcel.Quit();
        }
    }
}


