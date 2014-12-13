using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;

namespace docReader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
       public string templatePath = AppDomain.CurrentDomain.BaseDirectory + "Templates\\MyTemplate.docx";
       public string templatePathKont = AppDomain.CurrentDomain.BaseDirectory + "Templates\\MyTemplateKont.docx";
        public MainWindow()
        {
            InitializeComponent();
            ExcelPath.Text = AppDomain.CurrentDomain.BaseDirectory + "IVAN_BAZE_NEW_1.05.14.xls";

    
            
            DateTime time = DateTime.Now;
            var curMonth = time.ToString("MM");
            MontrhSelected.Text = curMonth.ToString();

           emailText.Text = "vidnessapsardze@gmail.com";
           emailPassword.Password = "redison003";
          // emailText.Text = "ivanb.dev@gmail.com";
          // emailPassword.Password = "K@nfetka";
            pathtoMail.Text = AppDomain.CurrentDomain.BaseDirectory + "PDFTemp";
        }
        Dictionary<string, string> mergeFields = new Dictionary<string,string>();
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
           // string excelPath = AppDomain.CurrentDomain.BaseDirectory + "List.xls"; //@"E:\sample.xlsx";
            if (!String.IsNullOrEmpty(ExcelPath.Text))
            {
                string konts = String.Empty;
                var source = SourceReader.TagsLibrary(ExcelPath.Text, MontrhSelected.Text);
                int countOfFiles = 0;
                docReader.ProgressDialog.ProgressDialogResult result = docReader.ProgressDialog.ProgressDialog.Execute(this.Owner, "Converting to Word", (bw, we) =>
                {
                foreach (var item in source)
                {
                    string value = item["PVN maks. nr. / personas kods:"].Trim();
                    string lig = item["Lig_nr"];
                    if (lig.Contains('/')) {
                        lig = lig.Replace('/', '_');
                    }
                    string mail = item["e-pasts"].Trim();
                    if (item.ContainsKey("konts"))
                    {
                         konts = item["konts"].Trim();
                        
                    }
                   // string templatePath = AppDomain.CurrentDomain.BaseDirectory + "Templates\\MyTemplate.docx";
                    string newDocPath =  value +"_"+lig;
                  
                    foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                    {
                        newDocPath = newDocPath.Replace(c, '_');
                    }
                    newDocPath = AppDomain.CurrentDomain.BaseDirectory + "Temp\\" + newDocPath + "[" + mail + "]" + ".docx";
                    if (!string.IsNullOrEmpty(konts))
                    {
                          try
                        {
                        File.Copy(templatePathKont, newDocPath);
                        konts = String.Empty;
                        }
                          catch (Exception ex)
                          {
                              MessageBox.Show(ex.Message.ToString());
                              Utilities.WriteLog(ex.Message.ToString());
                          }
                    }
                    else
                    {
                        try
                        {
                            File.Copy(templatePath, newDocPath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                            Utilities.WriteLog(ex.Message.ToString());
                        }
                    }
                     
                    ConvertWord.FillMergeFields(newDocPath, item);
                    countOfFiles++;
                    if (docReader.ProgressDialog.ProgressDialog.ReportWithCancellationCheck(bw, we, countOfFiles * 100 / source.Count(), "Executing step {0}/" + source.Count(), countOfFiles))
                                       return;
               
       
            }
                docReader.ProgressDialog.ProgressDialog.CheckForPendingCancellation(bw, we);
                },

           new docReader.ProgressDialog.ProgressDialogSettings(true, true, false));

                     if (result.Cancelled)
                         MessageBox.Show("ProgressDialog cancelled.");
                     else if (result.OperationFailed)
                         MessageBox.Show("ProgressDialog failed.");
                     else
                         MessageBox.Show("ProgressDialog successfully executed. " + countOfFiles + " files generated.", "Document Generation", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            else {
            
            }   
        }
        public void LoadExcel_Click(object sender, RoutedEventArgs e)
        {
            string filename = string.Empty;
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.DefaultExt = "*.xls; *.xlsx";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                // Open document
                 filename = dlg.FileName;
                ExcelPath.Text = filename;
            }    
        }
        private void ConvertToPDF(object sender, RoutedEventArgs e)
        {
            System.Windows.Window xxx = Application.Current.MainWindow;
            ConvertPDF.GeneratePdf(xxx);          
        }
        private void SendMails(object sender, RoutedEventArgs e)
        {
            System.Windows.Window xxx = Application.Current.MainWindow;
            SourceReader.SendMails(xxx, emailText.Text, emailPassword.Password, pathtoMail.Text, MontrhSelected.Text);
        }
    }
}
