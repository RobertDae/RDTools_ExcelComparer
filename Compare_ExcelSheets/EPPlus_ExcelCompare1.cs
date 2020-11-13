//using OfficeOpenXml;
using System;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using CsvHelper;
using CsvHelper.Configuration;
using EPPlus;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlus.Compare_ExcelSheets
{
    class EPPlus_ExcelCompare1
    {
        

        public EPPlus_ExcelCompare1()
        {

        }

        public EPPlus_ExcelCompare1(string file1, string file2, string outFile)
        {
            // copied code
            //check if compared files are csv then an import into datatables and a new excelfile is needed. 
            FileInfo fi1 = new FileInfo(file1);
            FileInfo fi2 = new FileInfo(file2);
            string ext1 = fi1.Extension;
            string ext2 = fi1.Extension;

            //check1
            DeleteOutputFileExists(outFile);

            //check2: if files do exist follow the normal procedure else throw error
            if (InputFilesExists(file1, file2))
            {
                
                //Check3: Translate csv into xlsx for better and standardized comparison
                if (ext1 == ".csv")
                {
                    DataTable dt = LoadCsv2Datatable(file1);
                    ImportToExcel(file1,dt);
                }
                if (ext2 == ".csv")
                {
                    DataTable dt = LoadCsv2Datatable(file2);
                    ImportToExcel(file2, dt);
                }
                // if not xlsx or csv files throw error and exit program
                if (((ext1 != ".xlsx") && (ext1 != ".csv")) || ((ext2 != ".xlsx") && (ext2 != ".csv")))
                {
                    Console.WriteLine("ERROR: Unsupported file extension! Please make sure you use .xlsx or .csv files as input.");
                    Environment.Exit(0);
                }
            }
            FileInfo plik1 = new FileInfo(file1);
            FileInfo plik2 = new FileInfo(file2);
            FileInfo path = new FileInfo(outFile);
                       

            using (ExcelPackage DiffResults = new ExcelPackage(path))
            using (ExcelPackage xlPackage = new ExcelPackage(plik1))
            using (ExcelPackage xlPackage2 = new ExcelPackage(plik2))
            {

                var worksheet1  = xlPackage.Workbook.Worksheets[0];
                var worksheet2 = xlPackage2.Workbook.Worksheets[0];
                var worksheet3 = DiffResults.Workbook.Worksheets.Add("ComparedDiffs");

                var maxCols = worksheet1.Cells.Columns;
                var maxRows = worksheet1.Cells.Rows;
                
                double diffCnt = 0;
                
                for (int row=1; row < (maxCols/100); row++)
                { 
                    for (int col=1; col < (maxRows/100); col++)
                    {
                        if ((worksheet1.Cells[row, col].Value != null) &&(worksheet2.Cells[row, col].Value!=null))  // if both are not null, we can convert them to strings
                            if (worksheet1.Cells[row, col].Value.ToString() != worksheet2.Cells[row, col].Value.ToString())
                            {
                                worksheet3.Cells[row,col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet3.Cells[row,col].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                worksheet3.Cells[row, col].Value = worksheet1.Cells[row, col].Value.ToString() + " --VS--> "+worksheet2.Cells[row, col].Value.ToString();
                                //Debug
                                //Console.WriteLine("Row:" +row+"/Col:"+col+", - values should be different:\n V1:\'"+ worksheet1.Cells[row, col].Value.ToString() + "\',\n V2:\'"+ worksheet2.Cells[row, col].Value.ToString()+"\'\n--------------------------------");
                                diffCnt++;
                            }
                    }

                }
                //CLI output: number of found diffs in both documents (int)
                Console.WriteLine(diffCnt);

                for(int i = 1; i <= worksheet3.Dimension.End.Column; i++)
                { 
                    worksheet3.Column(i).AutoFit();
                }
                worksheet3.Name = "DiffResults_TotalCnt_" + diffCnt;

                DiffResults.Save();
        }


        //end code

        }

        //private bool ConvertionCsvToExcelNeeded(ref FileInfo fi)
        //{
        //    bool IsCsv;
        //    Console.WriteLine("Debug: CSV-Import to Execl for " + fi.Name + ": ");
        //    //FileInfo fi_temp = ImportCsvToXlsx(ref fi);
        //    FileInfo convertedExcel = 
        //    //fi = fi_temp;
        //    IsCsv = true;
        //    return ;
        //}

        private bool InputFilesExists(string file1, string file2)
        {
            return (File.Exists(file1) && File.Exists(file2));
        }

        private void DeleteOutputFileExists(string outFile) 
        {
            if (File.Exists(outFile))
                 File.Delete(outFile);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fi1"> the name of the csv file, which needs to be ported into Excel </param>
        private void ImportToExcel(string XlsxFullname , DataTable dt)
        {
            //csv-Import
            //Source: https://riptutorial.com/de/epplus/example/26605/daten-aus-csv-datei-importieren
            //set the formatting options
            ExcelTextFormat format = new ExcelTextFormat();
            format.Delimiter = ';';
            format.Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString());
            format.Culture.DateTimeFormat.ShortDatePattern = "dd-mm-yyyy";
            format.Encoding = new UTF8Encoding();

            
            //create a new Excel package
            //string newFullFilename = fi1.DirectoryName + "\\" + Path.GetFileNameWithoutExtension(fi1.FullName.ToString()) + ".xlsx";
            
            FileInfo fiImported = new FileInfo(XlsxFullname);
           

            using (ExcelPackage excelPackage = new ExcelPackage(fiImported))
            {
                //excelPackage.SaveAs(fi)
                //create a WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("ImportedContentFromCSV");

                //load the CSV data into cell A1
                worksheet.Cells["A1"].LoadFromDataTable(dt,false); // load header set to false 
                excelPackage.Save();
            }
         
        }

        //private FileInfo ImportCsvToXlsx(ref FileInfo fi) 
        //{
        //    string csvFullFilename = fi.DirectoryName + "\\" + Path.GetFileName(fi.FullName.ToString());
        //    string xlsxFullFilename = fi.DirectoryName + "\\" + Path.GetFileNameWithoutExtension(fi.FullName.ToString()) + ".xlsx";

        //    using (var reader = new StreamReader(csvFullFilename))
        //    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        //    {
        //        // Do any configuration to `CsvReader` before creating CsvDataReader.
        //        using (var dr = new CsvDataReader(csv))
        //        {
        //            var dt = new DataTable();
        //            dt.Load(dr);
        //        }
        //    }


        //    return fi;
        //}

        private static DataTable LoadCsv2Datatable(string refPath)
        {
            var cfg = new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = ",", HasHeaderRecord = true };
            var result = new DataTable();
            using (var sr = new StreamReader(refPath, Encoding.UTF8, false, 16384 * 2))
            {
                using (var rdr = new CsvReader(sr, cfg))
                using (var dataRdr = new CsvDataReader(rdr))
                {
                    result.Load(dataRdr);
                }
            }
            return result;
        }

    }
}