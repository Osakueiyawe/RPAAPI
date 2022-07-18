using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace RPA_API.Methods
{
    public static class LogError
    {
        public static void Errhandler(string errormessage)
        {
            try
            {
                string loc = "C:/Error/" + DateTime.Today.ToString("dd-MM-yy");
                if (!Directory.Exists(loc))
                {
                    Directory.CreateDirectory(loc);
                }

                string path = loc + "/" + DateTime.Today.ToString("dd-MM-yy") + ".txt";
                if (!File.Exists(path))
                {
                    File.Create(path).Close();
                }
                using (StreamWriter w = File.AppendText(path))
                {
                    w.WriteLine("\r\nLog Entry : ");
                    w.WriteLine("{0}", DateTime.Now.ToString(CultureInfo.InvariantCulture));
                    //w.WriteLine(ui.customername); 
                    string err = "Error Message:" + errormessage;
                    w.WriteLine(err);
                    w.WriteLine("_______________________________________");
                    w.Flush();
                    w.Close();
                }
            }
            catch (Exception ex)
            {
                LogError.Errhandler(ex.Message);
            }
        }
}   }
