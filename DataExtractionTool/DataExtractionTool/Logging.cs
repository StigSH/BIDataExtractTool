using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataExtractionTool
{
    public class Logging
    {
        public static void WriteErrorToFile(string RootPth, Exception e)
        {
            

            DateTime now = DateTime.Now;
            string nowStr = now.ToString("yyyyMMdd_hh_mm_ss");
            string LogPath = RootPth + "\\Log";
            string LogFullFilePth = LogPath + "/Log" + nowStr + ".txt";
            string line = Environment.NewLine + Environment.NewLine;

            var st = new StackTrace(e, true);
            var frame = st.GetFrame(0);
            var lineNr = frame.GetFileLineNumber();

            try
            {
                if (!Directory.Exists(LogPath))
                {
                    Directory.CreateDirectory(LogPath);

                }


                using (StreamWriter sw = File.AppendText(LogFullFilePth))
                {
                    string error = "Log Written Date:" + " " + DateTime.Now.ToString() + line + "Error Line No :" + " " + lineNr + line + "Error Location : " + frame.GetFileName() + line + "Error Message:" + " " + e.Message + line + "Exception Type:" + " " + e.GetType() + line;
                    sw.WriteLine("-----------Exception Details on " + " " + DateTime.Now.ToString() + "-----------------");
                    sw.WriteLine("-------------------------------------------------------------------------------------");
                    sw.WriteLine(line);
                    sw.WriteLine(error);
                    sw.WriteLine("--------------------------------*End*------------------------------------------");
                    sw.WriteLine(line);
                    sw.Flush();
                    sw.Close();
                }
            }
            catch (Exception ErrorExcep)
            {
                Console.WriteLine(ErrorExcep.Message);
            }
        }
    }
}
