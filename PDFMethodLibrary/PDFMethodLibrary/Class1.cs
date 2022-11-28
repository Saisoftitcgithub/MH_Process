using IronOcr;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
//using System.Reflection.Metadata;
//using System.Reflection.PortableExecutable;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;


namespace PDFMethodLibrary
{
    public class PDFClass
    {
        public void ExtractPDFContent(string pdfPath)
        {
            var Ocr = new IronTesseract();
            using (var Input = new OcrInput())
            {
                //read pdf 
                Input.AddPdf(pdfPath);
                var Result = Ocr.Read(Input);
                Result.SaveAsTextFile("Result.txt");
            }
        }

        public DataTable ExtractPDFValues()
        {
            //create datatable
            DataTable dt = new DataTable();
            DataRow row;
            dt.Columns.Add("Page Number", typeof(String));
            dt.Columns.Add("Tax Invoice Number", typeof(String));
            dt.Columns.Add("Page Count", typeof(String));

            int linenum = 0;
            row = dt.NewRow();
            foreach (string line in File.ReadLines("Result.txt"))
            {
                string pageNum;
                string taxInvoiceNum;

                string temp_line = line.Replace(" ", "").ToLower();

                //Console.WriteLine(temp_line.Length);

                if (temp_line.Length != 0)
                {
                    //row = dt.NewRow();
                    if (temp_line.Contains("pageno:"))
                    {
                        pageNum = temp_line.Substring(temp_line.IndexOf("pageno:") + 7, (temp_line.Length - ((temp_line.IndexOf("pageno:")) + 7)));
                        if (pageNum.Length > 6)
                        {
                            string line2 = line.ToLower();
                            string pageNum2 = line2.Substring(line2.IndexOf("page no") + 7, (line2.Length - ((line2.IndexOf("page no")) + 7)));

                            pageNum2 = pageNum2.Replace(":", "");
                            int indexOfLastSpace = pageNum2.LastIndexOf(' ');
                            pageNum2 = pageNum2.Substring(0, indexOfLastSpace);
                            pageNum = pageNum2.Replace(" ", "");

                        }
                        row["Page Number"] = pageNum;

                        //Extract total page count
                        String Num_Pages_temp = pageNum.Substring(pageNum.LastIndexOf('f') + 1);
                        String Num_Pages;
                        if (pageNum.Substring(pageNum.LastIndexOf('f') + 1).Trim() == "it")
                        {
                            Num_Pages = "11";
                        }
                        else
                        {
                            Num_Pages = Num_Pages_temp;
                        }
                        row["Page Count"] = Num_Pages;

                    }
                    else if (temp_line.Contains("no:"))
                    {
                        pageNum = temp_line.Substring(temp_line.IndexOf("no:") + 3, (temp_line.Length - ((temp_line.IndexOf("no:")) + 3)));
                        if (pageNum.Length > 6)
                        {
                            string line2 = line.ToLower();
                            string pageNum2 = line2.Substring(line2.IndexOf("no") + 2, (line2.Length - ((line2.IndexOf("no")) + 2)));

                            pageNum2 = pageNum2.Replace(":", "");
                            int indexOfLastSpace = pageNum2.LastIndexOf(' ');
                            pageNum2 = pageNum2.Substring(0, indexOfLastSpace);
                            pageNum = pageNum2.Replace(" ", "");

                        }
                        row["Page Number"] = pageNum;

                        //Extract total page count
                        String Num_Pages_temp = pageNum.Substring(pageNum.LastIndexOf('f') + 1);
                        String Num_Pages;
                        if (pageNum.Substring(pageNum.LastIndexOf('f') + 1).Trim() == "it")
                        {
                            Num_Pages = "11";
                        }
                        else
                        {
                            Num_Pages = Num_Pages_temp;
                        }
                        row["Page Count"] = Num_Pages;

                    }


                    if (temp_line.Contains("taxinvoiceno"))
                    {
                        taxInvoiceNum = temp_line.Substring(temp_line.IndexOf("taxinvoiceno") + 12, (temp_line.Length - ((temp_line.IndexOf("taxinvoiceno")) + 12)));

                        row["Tax Invoice Number"] = taxInvoiceNum;
                        dt.Rows.Add(row);
                        row = dt.NewRow();

                    }
                    else if (temp_line.Contains("taxinveiceno"))
                    {
                        taxInvoiceNum = temp_line.Substring(temp_line.IndexOf("taxinveiceno") + 12, (temp_line.Length - ((temp_line.IndexOf("taxinveiceno")) + 12)));

                        row["Tax Invoice Number"] = taxInvoiceNum;
                        dt.Rows.Add(row);
                        row = dt.NewRow();
                    }

                }

                linenum++;

            }
            Console.WriteLine(dt.Rows.Count);

            //print data rows
            /*
            foreach (DataRow row1 in dt.Rows)
            {
                string pgno = row1["Page Number"].ToString();
                string taxno = row1["Tax Invoice Number"].ToString();
                string noofpages = row1["Page Count"].ToString();
                Console.WriteLine(pgno + " " + taxno + " " + noofpages);

            }*/

            //export to excel
            var lines = new List<string>();

            string[] columnNames = dt.Columns
                .Cast<DataColumn>()
                .Select(column => column.ColumnName)
                .ToArray();

            var header = string.Join(",", columnNames.Select(name => $"\"{name}\""));
            lines.Add(header);

            var valueLines = dt.AsEnumerable()

                .Select(row => string.Join(",", row.ItemArray.Select(val => $"\"{val}\"")));

            lines.AddRange(valueLines);

            File.WriteAllLines("excel.csv", lines);

            return dt;

        }

        public void ExtractRange_Split_Compress(DataTable dt, string pdfPath, string opPath, string cpPath)
        {
            //Print Data table
            /*
            foreach (DataRow row1 in dt.Rows)
            {
                string pgno = row1["Page Number"].ToString();
                string taxno = row1["Tax Invoice Number"].ToString();
                string noofpages = row1["Page Count"].ToString();
               // Console.WriteLine(pgno + " " + taxno + " " + noofpages);

            }    */

            int counter = 0;
            int counter_To = 0;
            int subPageCount = 0;
            int rangeFrom = 0;
            int rangeTo = 0;
            foreach (DataRow row1 in dt.Rows)
            {
                counter = counter + 1;
                String noofpages = row1["Page Count"].ToString();
                subPageCount = Int32.Parse(noofpages);
                //Console.WriteLine(subPageCount);
                counter_To = rangeTo + subPageCount;
                //Console.WriteLine(Counter + "-" + Counter_To);
                if (counter == counter_To)
                {
                    rangeFrom = rangeFrom + 1;
                    rangeTo = counter_To;
                    String taxInvNo = row1["Tax Invoice Number"].ToString();
                    Console.WriteLine(rangeFrom + "-" + rangeTo + "-" + taxInvNo);
                    PDFClass p1 = new PDFClass();
                    p1.splitAndCompressPDF(pdfPath, opPath, cpPath, rangeFrom, rangeTo, taxInvNo);
                    rangeFrom = rangeFrom + subPageCount - 1;
                }

            }

            //Read Excel Data
            /*
            var path = "excel.csv";
            using (TextFieldParser csvParser = new TextFieldParser(path))
            {
                csvParser.CommentTokens = new string[] { "#" };
                csvParser.SetDelimiters(new string[] { "," });
                csvParser.HasFieldsEnclosedInQuotes = true;

                // Skip the row with the column names
                csvParser.ReadLine();

                while (!csvParser.EndOfData)
                {
                    // Read current line fields, pointer moves to the next line.
                    string[] fields = csvParser.ReadFields();
                    string PageNo = fields[0];
                    string InvNo = fields[1];

                }
            }*/

        }

        private void splitAndCompressPDF(string pdfPath, string opPath, string cpPath, int rangeFrom, int rangeTo, string taxInvNo)
        {
            using (PdfReader reader = new PdfReader(pdfPath))
            {
                Document document = new Document();
                PdfCopy copy = new PdfCopy(document, new FileStream(opPath + "\\" + taxInvNo + ".pdf", FileMode.Create));
                document.Open();
                for (int pagenumber = rangeFrom; pagenumber <= rangeTo; pagenumber++)
                {
                    if (reader.NumberOfPages >= pagenumber)
                    {
                        copy.AddPage(copy.GetImportedPage(reader, pagenumber));
                    }
                    else
                    {
                        break;
                    }

                }
                //PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);
                //pdfRenderer.PdfDocument = new PdfSharp.Pdf.PdfDocument();
                //pdfRenderer.PdfDocument.Options.FlateEncodeMode = PdfFlateEncodeMode.BestCompression;
                document.Close();
                //code for compress pdf

                //String cpPath = @"C:\Office\Soujanya\VisualStudio\MH_DOCS\OUTPUT\TAXINVOICE\TI_COMPRESSED";

                String compressPDFInputFileName = opPath + "\\" + taxInvNo + ".pdf";
                String compressPDFOutputFileName = cpPath + "\\" + taxInvNo + "_Compressed" + ".pdf";

                try
                {

                    // string command = "gs -dNOPAUSE -sDEVICE=pdfwrite -dBATCH -dPDFSETTINGS=/screen" +
                    //  "-sOutputFile=\" C:\\Office\\Soujanya\" \" C:\\Office\\Soujanya\\TAXINVOICE_organized.pdf";

                    string command = $"gswin64c -sDEVICE=pdfwrite -dNOPAUSE -dBATCH -q -dPDFSETTINGS=/screen -o {compressPDFOutputFileName} {compressPDFInputFileName}";

                    // string command = "gswin64c -sDEVICE=pdfwrite -dNOPAUSE -dBATCH -q -dPDFSETTINGS=/screen -o "+compressPDFOutputFileName+" "+compressPDFInputFileName+"";

                    Process pdfProcess = new Process();

                    StreamWriter writer;
                    StreamReader reader1;

                    ProcessStartInfo info = new ProcessStartInfo("cmd.exe", "/K " + command);
                    info.WorkingDirectory = System.AppDomain.CurrentDomain.BaseDirectory;

                    info.CreateNoWindow = true;
                    info.UseShellExecute = false;
                    info.RedirectStandardInput = true;
                    info.RedirectStandardOutput = true;

                    pdfProcess.StartInfo = info;
                    pdfProcess.Start();

                    writer = pdfProcess.StandardInput;
                    reader1 = pdfProcess.StandardOutput;
                    writer.AutoFlush = true;

                    writer.WriteLine(command);

                    writer.Close();

                    string ret = reader1.ReadToEnd();
                }
                catch (Exception ex)
                {
                    throw ex;
                }


            }
        }

    }
}