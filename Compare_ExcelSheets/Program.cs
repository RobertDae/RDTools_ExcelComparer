using EPPlus.Compare_ExcelSheets;
using System;
 

namespace Compare_ExcelSheets
{
    class Program
    {
        static void Main(string[] args)
        {
            //incase not all arguments are correctly filled, display error message on CLI
            if (args.Length != 3)
            {
                Console.WriteLine("ERROR: To few agruments please check arguments:\nCompare_Excelsheets <Fullpath_of_file1_to_Comapare> <Fullpath_of _file2_to_comapare> <Fullpath_outputfile>");
                Environment.Exit(0);
            }
            //if all fits 
            else
            {
                EPPlus_ExcelCompare1 ExcelCompObj = new EPPlus_ExcelCompare1(args[0], args[1], args[2]);
            }
        }
    }
}
