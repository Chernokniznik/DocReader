using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace docReader
{
  public  class Utilities
    {
        public static void WriteLog(String message)
        
       {
            if (String.IsNullOrEmpty(message))
                return;
            try
            {
                String curentDir = String.Format("{0}\\Logs\\", Path.GetDirectoryName(Assembly.GetEntryAssembly().Location));
                if (!Directory.Exists(curentDir))
                    Directory.CreateDirectory(curentDir);

                String logName = String.Format("{0}\\{1}{2}{3}.log",
                   curentDir,
                    DateTime.Today.Year.ToString(),
                    DateTime.Today.Month.ToString("00"),
                    DateTime.Today.Day.ToString("00"));

                String formatedMessage = String.Format("[{0}] \t{1}", DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"),
                     message);

                // parameter append = true, allows to create file if not found				
                using (StreamWriter stream = new System.IO.StreamWriter(logName, true))
                {
                    stream.WriteLine(formatedMessage);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Error writing to the log file: {0}", ex.Message));
            }
        }
       
    }
   

}
