using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Windows;
using System.Linq;
namespace docReader
{
   public static class SourceReader
    {
       public static string pathToexcel { get; set; }
       public  static string htmlFilePath = AppDomain.CurrentDomain.BaseDirectory + "mail.txt";
       public static List<Dictionary<string, string>> TagsLibrary(string path, string selectedMonth)
       {
           int selectedMonthInt = 0;
           int.TryParse(selectedMonth, out selectedMonthInt);
           pathToexcel = path;
           List<Dictionary<string, string>> invoiceList = new List<Dictionary<string, string>>();
           Dictionary<string, string> invoice = new Dictionary<string, string>();
           FileStream stream = null;
           IExcelDataReader excelReader = null;
           string extension = String.Empty;

           if (!string.IsNullOrEmpty(path) && selectedMonthInt > 0)
           {
               extension = path.Substring(path.Length - 4);
               try
               {
                   stream = File.Open(path, FileMode.Open, FileAccess.Read);
                   switch (extension)
                   {
                       case ".xls":
                           excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                           break;
                       case "xlsx":
                           excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                           break;
                   }
               }
               catch (Exception ex)
               {
                   Utilities.WriteLog(ex.ToString());
                   MessageBox.Show("Excel file can't be executed, please close excel file");
               }

               if (excelReader != null)
               {
                   excelReader.IsFirstRowAsColumnNames = true;
                   DataSet result = excelReader.AsDataSet();
                   DataTable table = result.Tables[0];

                   var nonEmptyRows = (from DataRow record in result.Tables[0].AsEnumerable()
                                       where !string.IsNullOrEmpty(record[0].ToString())
                                       select record).Count();

                   for (int i = 0; nonEmptyRows > i; i++)
                   {
                       invoice = new Dictionary<string, string>();
                       invoice.Add("Lig_nr", table.Rows[i].ItemArray[0].ToString());
                       invoice.Add("Maksatajs", table.Rows[i].ItemArray[1].ToString());
                       invoice.Add("PVN maks. nr. / personas kods:", table.Rows[i].ItemArray[2].ToString());
                       invoice.Add("Objekta adrese", table.Rows[i].ItemArray[3].ToString());
                       invoice.Add("Juridiskā adrese / Dekl.dz.v.", table.Rows[i].ItemArray[4].ToString());
                       invoice.Add("Zvans", table.Rows[i].ItemArray[5].ToString());
                       invoice.Add("01.2014", table.Rows[i].ItemArray[8].ToString());
                       invoice.Add("02.2014", table.Rows[i].ItemArray[9].ToString());
                       invoice.Add("03.2014", table.Rows[i].ItemArray[10].ToString());
                       invoice.Add("04.2014", table.Rows[i].ItemArray[11].ToString());
                       invoice.Add("05.2014", table.Rows[i].ItemArray[12].ToString());
                       invoice.Add("06.2014", table.Rows[i].ItemArray[13].ToString());
                       invoice.Add("07.2014", table.Rows[i].ItemArray[14].ToString());
                       invoice.Add("08.2014", table.Rows[i].ItemArray[15].ToString());
                       invoice.Add("09.2014", table.Rows[i].ItemArray[16].ToString());
                       invoice.Add("10.2014", table.Rows[i].ItemArray[17].ToString());
                       invoice.Add("11.2014", table.Rows[i].ItemArray[18].ToString());
                       invoice.Add("12.2014", table.Rows[i].ItemArray[19].ToString());
                       invoice.Add("kontakti", table.Rows[i].ItemArray[24].ToString());
                       invoice.Add("e-pasts", table.Rows[i].ItemArray[20].ToString());
                       float tmpGParads = 0.00f;
                       float kopaEUR = 0.00f;

                       float.TryParse(table.Rows[i].ItemArray[7].ToString(), out kopaEUR);
                       if (kopaEUR > 0)
                       {
                           invoice.Add("abon", kopaEUR.ToString("f2"));
                       }
                       if (table.Rows[i].ItemArray.Length > 21 && !string.IsNullOrEmpty(table.Rows[i].ItemArray[21].ToString()))
                       {
                           float.TryParse(table.Rows[i].ItemArray[21].ToString(), out tmpGParads);
                       }
                       if (table.Rows[i].ItemArray.Length > 22 && !string.IsNullOrEmpty(table.Rows[i].ItemArray[22].ToString()))
                       {
                           invoice.Add("konts", table.Rows[i].ItemArray[22].ToString());
                       }

                       if (table.Rows[i].ItemArray.Length > 23 && !string.IsNullOrEmpty(table.Rows[i].ItemArray[23].ToString()))
                       {
                           invoice.Add("banka", table.Rows[i].ItemArray[23].ToString());
                       }
                       #region variable
                       float parads = 0.0000f;
                       float pvn = 0.0000f;
                       float summa_bezPvn = 0.0000f;
                       float minusCurMonth = 0.0000f;
                       float kopa = 0.0000f;
                       float minusCurMonthParads = 0.000f;
                       DateTime date = DateTime.Now;
                       var month = new DateTime(date.Year, selectedMonthInt, 1);
                       var first = month;
                       var last = month.AddMonths(1).AddDays(-1);
                       string cur_Year = date.ToString("yyyy");
                       string date_ToPay = String.Empty;                 
                       #endregion
                       invoice.Add("generated_date", month.ToString("dd.MM.yyyy"));
                       date_ToPay = month.AddDays(9.0).ToString("dd.MM.yyyy");
                       invoice.Add("date_toPay", date_ToPay);                  
                       invoice.Add("cur_month", selectedMonth + cur_Year.Substring(2));
                       invoice.Add("first_day", first.ToString("dd.MM.yyyy"));
                       invoice.Add("last_day", last.ToString("dd.MM.yyyy"));
                       int cnt = 0;
                       for (int m = 8; m <= 8+selectedMonthInt; m++)
                       {
                           float tmpmoney = 0.0000f;
                           string tmpValue = string.Empty;
                           tmpValue = table.Rows[i].ItemArray[m].ToString();
                           if (!string.IsNullOrWhiteSpace(tmpValue) && tmpValue.Contains(','))
                           {
                               tmpValue = tmpValue.Replace(',', '.');
                           }
                           float.TryParse(tmpValue, out tmpmoney);
                           if (tmpmoney > 0)
                           {
                               cnt++;
                           }
                           parads += tmpmoney;
                       }
                       var curMonthValue = invoice.First(x => x.Key.Equals((selectedMonth +"."+ cur_Year).ToString())).Value;
                       float.TryParse(curMonthValue.ToString(), out minusCurMonth);
                       float minus = 0.0000f;
                       if (minusCurMonth >= 0)
                       {
                           minusCurMonthParads = parads - minusCurMonth;
                       }
                       if (parads < 0 || minusCurMonthParads < 0)
                       {
                           minus = minusCurMonthParads;
                           minusCurMonthParads = 0;
                       }
                       invoice.Add("parads", minusCurMonthParads.ToString("f2"));
                       kopa = tmpGParads + minusCurMonthParads;
                       if (kopa > 0)
                       {
                           invoice.Add("kopa", (kopa).ToString("f2"));
                           pvn = ((kopa) / 1.21f) * 0.21f;
                           summa_bezPvn = (kopa) - pvn;
                           invoice.Add("pvn", pvn.ToString("f2"));
                           invoice.Add("summa_bezPvn", ((kopa) - pvn).ToString("f2"));
                           var numberToWords = NumberToWords.NumberToWord((kopa).ToString("f2"));
                           invoice.Add("numbersToWords", numberToWords);
                           if (tmpGParads > 0)
                           {
                               invoice.Add("gparads", tmpGParads.ToString("f2"));
                           }
                           else
                           {
                               int y = 0;
                               invoice.Add("gparads", y.ToString("f2"));
                           }
                       }
                       else
                       {
                           kopa = 0;
                           invoice.Add("kopa", (kopa).ToString("f2"));
                           invoice.Add("pvn", pvn.ToString("f2"));
                           invoice.Add("summa_bezPvn", ((kopa)).ToString("f2"));
                           var numberToWords = NumberToWords.NumberToWord((kopa).ToString("f2"));
                           invoice.Add("numbersToWords", numberToWords);
                           invoice.Add("gparads", kopa.ToString("f2"));
                       }

                       if (minusCurMonth > 0)
                       {
                           if (minus < 0)
                           {
                               minusCurMonth += minus;
                           }
                           invoice.Add("cur_sum", minusCurMonth.ToString("f2"));
                           float tmpCur_sumPVN = 0.0000f;
                           tmpCur_sumPVN = ((minusCurMonth) / 1.21f) * 0.21f;
                           invoice.Add("ecur_sum", (minusCurMonth - tmpCur_sumPVN).ToString("f2"));
                           invoice.Add("epvn", tmpCur_sumPVN.ToString("f2"));
                           var numberToWords1 = NumberToWords.NumberToWord((minusCurMonth).ToString("f2"));
                           invoice.Add("numbersToWordsCur", numberToWords1);
                       }
                       else
                       {
                           minusCurMonth = 0;
                           invoice.Add("cur_sum", minusCurMonth.ToString("f2"));
                           invoice.Add("ecur_sum", minusCurMonth.ToString("f2"));
                           invoice.Add("epvn", minusCurMonth.ToString("f2"));
                           var numberToWords1 = NumberToWords.NumberToWord((minusCurMonth).ToString("f2"));
                           invoice.Add("numbersToWordsCur", numberToWords1);
                       }

                       if (cnt > 0 && minusCurMonth > 0)
                       {
                           cnt = cnt - 1;
                       }
                       invoice.Add("m_cnt", (cnt).ToString());
                       float all_eiro = 0.0000f;
                       all_eiro = kopa + minusCurMonth;
                       invoice.Add("all_eiro", all_eiro.ToString("f2"));
                       var all_eiro_string = NumberToWords.NumberToWord((all_eiro).ToString("f2"));
                       invoice.Add("all_eiro_words", all_eiro_string);
                       invoiceList.Add(invoice);
                   }
                   excelReader.Close();
               }
           }
           return invoiceList;
       }

       public static void SendMails(Window owner, string email, string emailPassword, string pathToSend, string selectedMonth)
       {
 
           if (!string.IsNullOrEmpty(email) && !string.IsNullOrEmpty(emailPassword) && !string.IsNullOrEmpty(pathToSend))
           { 
           DirectoryInfo di = new DirectoryInfo(/*AppDomain.CurrentDomain.BaseDirectory + "Temp"*/ pathToSend);       
           var directories = di.GetFiles("*", SearchOption.AllDirectories);
           int countOfFiles = 0;

           docReader.ProgressDialog.ProgressDialogResult result = docReader.ProgressDialog.ProgressDialog.Execute(owner, "Sending mails", (bw, we) =>
           {
           foreach (FileInfo d in directories)
           {
               int recCount = ReturnMailsCount(directories);
               string s = d.Name;
               if (d.Name.Length > 7 && !d.Name.Contains('~'))
               {
                   if (htmlFilePath != null)
                   {
                       using (StreamReader reader = File.OpenText(htmlFilePath)) // Path to your 
                       {
                           DateTime date = DateTime.Now;
                           string curY = date.ToString("yyyy");
                           string tmp = String.Empty;
                           string Mainresult;
                           string[] item;
                           NewMethod(s, out Mainresult, out item);
                           if (Mainresult.Contains('@'))
                           {
                               foreach (var t in item)
                               {
                                   tmp = SendSimpleMail(email, emailPassword, pathToSend, ref countOfFiles, d, reader, selectedMonth +"."+curY, t.Trim());

                               }
                               if (String.IsNullOrEmpty(tmp))
                               {
                                   MoveFile(pathToSend + "\\" + d.Name, AppDomain.CurrentDomain.BaseDirectory + "SendMails\\" + d.Name);
                                   if (docReader.ProgressDialog.ProgressDialog.ReportWithCancellationCheck(bw, we, countOfFiles * 100 / recCount, "Executing step {0}/" + recCount, countOfFiles))
                                       return;
                               }
                           }
                       }
                   }
                   else
                   {
                       MessageBox.Show("Mail template not exist !");
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
                  MessageBox.Show("ProgressDialog successfully executed. " + countOfFiles + " files send.", "Mails Send", MessageBoxButton.OK, MessageBoxImage.Information);
        
           }
           else
           {
               MessageBox.Show("something is empty!");
           }         
       }

       private static string SendSimpleMail(string email, string emailPassword, string pathToSend, ref int countOfFiles, FileInfo d, StreamReader reader, string curMonth, string t)
       {
           string tmp = String.Empty;
           if (!string.IsNullOrEmpty(t) && t.Contains('@'))
           {
               countOfFiles++;
               //var x = "elena.sterehova@gmail.com";
               //x = "vidness@inbox.lv";
               SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587);
               NetworkCredential basicCredential = new NetworkCredential(email, emailPassword);
               MailMessage message = new MailMessage();
               MailAddress fromAddress = new MailAddress(email);
               // setup up the host, increase the timeout to 5 minutes
               smtpClient.UseDefaultCredentials = false;
               smtpClient.EnableSsl = true;
               smtpClient.Credentials = basicCredential;
               smtpClient.Timeout = (60 * 5 * 1000);
               message.From = fromAddress;
               message.Subject = "Faktūrrēķins SIA Vidness par " + curMonth;
               message.IsBodyHtml = true;
               message.Body = reader.ReadToEnd();
               message.To.Add(t);
               string attachmentFilename = pathToSend + "\\" + d.Name;
               if (attachmentFilename != null)
               {
                   try
                   {
                       message.Attachments.Add(new Attachment(attachmentFilename));
                   }
                   catch (Exception ex)
                   {
                       MessageBox.Show(ex.ToString());
                       Utilities.WriteLog(ex.ToString());
                   }
               }
               try
               {
                     smtpClient.Send(message);
               }
               catch (Exception ex)
               {
                   Utilities.WriteLog(ex.ToString());
                   tmp = ex.ToString();
               }
               smtpClient.Dispose();
               message.Dispose();
               GC.Collect();
               GC.WaitForPendingFinalizers();
           }
           return tmp;
       }

       private static void NewMethod(string s, out string Mainresult, out string[] item)
       {
           int start = s.IndexOf("[") + 1;
           int end = s.IndexOf("]", start);
           Mainresult = s.Substring(start, end - start);
           item = Mainresult.Split(';');
       }

       private static int ReturnMailsCount(FileInfo[] directories)
       {
           int cnt = 0;

           foreach (var d in directories)
           {
               //cnt = directories.Count();
               string s = d.Name;
               string Mainresult;
               string[] item;
               NewMethod(s, out Mainresult, out item);
               foreach (var t in item) {
                  if(!string.IsNullOrEmpty(t) && t.Contains('@')){
                   cnt++;
                  }
               }
           }           
           return cnt;
       }
       public static bool MoveFile(string sourceFile, string destinationFile)
       {
           bool result = false;

           if (File.Exists(sourceFile) == true && !File.Exists(destinationFile))
               try
               {
                   File.Move(sourceFile, destinationFile);
                   result = true;
               }
               catch (Exception e)
               {
                   Utilities.WriteLog(e.ToString());
               }
           return result;
       }
    }
}
