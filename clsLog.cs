using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


//Version Control
// Who When        Ref Notes
// MMF 09-Mar-2013 --- New

namespace onelan
{
    class clsLog
    {
        string msVersionData, msLogFile;

        //constructor
        public clsLog(string sLogFile, string sVersionData)
        {
            //set log filepath and version data
            msLogFile = sLogFile;
            msVersionData = sVersionData;
        }

        public bool mLog(string sStatus,string sMessage)
        {
            //determine the Log text
            string sLogText;
            
            //log text to be <Date(DD-MMM-YYYY HH:NN:SS)>; <VersionNo> <VersionDate>; <Username>; <Status>; <custom input>
            //Add generic Log message info
            sLogText = DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss") + "; " + msVersionData + ", " + Environment.UserName + "; " + sStatus + "; ";

            //add the message for hte Log
            sLogText = sLogText + sMessage;

            //write to the Log file
            StreamWriter sw = File.AppendText(msLogFile);
            sw.WriteLine(sLogText);
            sw.Close();

            return true;
        }
    }
}
