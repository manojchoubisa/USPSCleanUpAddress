using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace USPSCleanUp
{
    public static class Logger
    {
        public static void Log(Exception Ex, string statusMessgae)
        {
            string logPath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            string FileName = "Logger.txt";
            string fullLogPAth = logPath + "\\" + FileName;

            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            StreamWriter fs = null;
            try
            {
                using (fs = File.AppendText(fullLogPAth))

                //using (fs = File.OpenWrite(fullLogPAth))
                {
                    if (Ex != null)
                    {
                        string message = "\r\nException : " + DateTime.Now.ToString() + " => " + Ex.Message;

                        fs.WriteLine(message);

                        message = "\nStackTrace : " + DateTime.Now.ToString() + " => " + Ex.StackTrace;

                        fs.WriteLine(message);
                        fs.WriteLine("-------------------------------");
                    }

                    if (!string.IsNullOrEmpty(statusMessgae))
                    {
                        string message = "\r\nMessgae : " + DateTime.Now.ToString() + " => " + statusMessgae;
                        fs.WriteLine(message);
                        fs.WriteLine("-------------------------------");
                    }

                    fs.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in Logger : " + ex.Message);
            }
            finally
            {
                if (fs != null)
                    fs.Close();
            }

        }
    }
}
