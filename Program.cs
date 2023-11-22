using ExtractExcelApp.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractExcelApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var enviroment = System.Environment.CurrentDirectory;
                string projectDirectory = Directory.GetParent(enviroment).Parent.FullName;

                var inputFile = projectDirectory + "\\Input\\Bayer Accounts.xlsx";

                ExportExccl exportExcel = new ExportExccl();
                exportExcel.ExportDataAndWriteToText(inputFile);

            }
            catch (Exception ex )
            {
                throw ex;
            }
            Console.ReadLine();
        }
    }
}