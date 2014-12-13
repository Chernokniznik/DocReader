using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace docReader
{
    public class ConvertPDF
    {

        public static void GeneratePdf(System.Windows.Window owner)
        {
            DirectoryInfo di = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "Temp");
            var directories = di.GetFiles("*", SearchOption.AllDirectories);
            int recCount = ReturnMailsCount(directories);
            int countOfFiles =0;
             docReader.ProgressDialog.ProgressDialogResult result = docReader.ProgressDialog.ProgressDialog.Execute(owner, "Converting to PDF", (bw, we) =>
           {
            foreach (FileInfo d in directories)
            {
                if (d.Name.Length > 7 && !d.Name.Contains('~'))
                {

                    try
                    {
                        Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
                        var wordDocument = appWord.Documents.Open(AppDomain.CurrentDomain.BaseDirectory + "Temp\\" + d.Name);
                        var minus = d.Name.Substring(0, d.Name.Length - 4).ToString();
                        wordDocument.ExportAsFixedFormat(AppDomain.CurrentDomain.BaseDirectory + "PDFtemp\\" + minus.ToString() + "pdf", WdExportFormat.wdExportFormatPDF);
                        wordDocument.Close();
                        appWord.Quit();
                        wordDocument = null;
                        appWord = null;

                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        countOfFiles++;
                         if (docReader.ProgressDialog.ProgressDialog.ReportWithCancellationCheck(bw, we, countOfFiles * 100 / recCount, "Executing step {0}/" + recCount, countOfFiles))
                                   return;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Word not found!");
                        Utilities.WriteLog(ex.ToString());
                    }
                }
            }
            docReader.ProgressDialog.ProgressDialog.CheckForPendingCancellation(bw, we);
           },

           new docReader.ProgressDialog.ProgressDialogSettings(true, true, false));

             if (result.Cancelled)
                 MessageBox.Show("ProgressDialog cancelled.");
             else if (result.OperationFailed)
                 MessageBox.Show("ProgressDialog failed.");
             else
                 MessageBox.Show("ProgressDialog successfully executed. " + countOfFiles + " files generated.", "PDF generation", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        private static int ReturnMailsCount(FileInfo[] directories)
        {
            int cnt = 0;
            foreach (var d in directories)
            {
                if (d.Name.Length > 7 && !d.Name.Contains('~'))
                {
                    cnt++;
                }
            }
            return cnt;
        }


    }
}
