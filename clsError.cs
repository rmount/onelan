using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


//Version Control
// Who When        Ref Notes
// NRE 23-Mar-2013 --- Enhance MessgeBox
// MM  26-Mar-2013     Added Summary data
// NRE 09-Jun-2014 1080 Allow blocking of message box of error
// NRE 27-Sep-2015      Log to console

namespace onelan
{
    class clsError
    {
        /// <summary>
        /// Display an Error Message to the user and Log the Error.
        /// </summary>
        /// <param name="sInfo">Text to explain the error to the user.</param>
        /// <param name="sClass">Class the error was encountered in.</param>
        /// <param name="sMethod">Method the error was encountered in.</param>
        /// <param name="e">Exception.</param>
        /// <param name="sVersionData">Version Number and Published Date.</param>
        /// <param name="sLogFile">Fileapth to the Log file to write the error to.</param>
        /// <returns></returns>
        public bool mLogError(string sInfo, string sClass, string sMethod, Exception e, string sVersionData, string sLogFile, Boolean bMsgbox = true)
        {
            string sErrorMessage;
            
            //save the error message
            sErrorMessage = sInfo + "; Class = " + sClass + "; Method = " + sMethod + "; Error = " + e.HResult + "; Description = " + e.Message;

            //display error message to the user
           // MessageBox.Show(sErrorMessage);
           if (bMsgbox) {
               MessageBox.Show(sErrorMessage, Constants.gcAppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
           }

            //initialise an instance of the log class
            clsLog log = new clsLog(sLogFile, sVersionData);

            //send the error message to be logged
            log.mLog(Constants.gcError, sErrorMessage);
            Console.WriteLine(sErrorMessage);

            return true;
        }
    }
}
