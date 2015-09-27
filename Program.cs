using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using WinSCP;


// Program onelan
// Author : NRE
// Date : 26-Sep-2015
// Purpose : Onelan interface in c#
// See 1542
// Note : This version does not attempt any file compression

// Methods:

// public static Boolean bProcessImage(String sSourceFolder, String sSourceImage, String sNewFolder)
// public static Boolean bRefreshData(String sDatabaseLocation)
// public static Boolean bExportDataToXML(String sDatabaseLocation)
// public static Boolean bsFtp(String sFtpSite, String sUserName, String sPassword, String sSshHostKeyFingerprint, String sDatabaseLocation, Boolean bSearchAll = false, Boolean bDebug = false, Boolean bExposeCredentials = false)
 


namespace onelan
{
    public class onelan
    {
        // Name of processing db. This has connections to PropData and the MySql database
        public const String gsOneLanDatabase = "onelan.mdb";
        public const String gcDSN = "Provider=Microsoft.ACE.OLEDB.12.0.0;Data Source=";

        // Indent for debug logs
        public const String sIndent = "  ";

        public const String msLogFile= "M:\\Property\\onelan\\onelan.log";
        public const String msVersionData = Constants.gcVersion + "; " + Constants.gcVersionDate;

        // Alert user of errors?
        public const Boolean bAlertUser = false;

        static void Main(string[] args)
        {
            String sSourceDir = null;
            String sTargetDir = null;
            String sDatabaseLocation = null;
            String sNewFolder = null;
	        String sFtpSite = null;
	        String sFtpSiteUserName = null;
	        String sFtpSitePassword = null;
	        String sHostKey = null;
	        String  sDebug = null;
            Boolean bDebug = false;

            String sPostToSftp = null;
            Boolean bPostToSftp = false;

            Boolean bOk = false;

            clsError clsError_ = new clsError();
            clsLog clsLog_ = new clsLog(msLogFile, msVersionData);


            try
            {

                clsLog_.mLog(Constants.gcInfo, "");
                clsLog_.mLog(Constants.gcInfo, "");
                clsLog_.mLog(Constants.gcInfo, "*** Starting RC Onelan Interface ...");

                // Collect settings
                sSourceDir = ConfigurationManager.AppSettings["SourceDir"];
                sTargetDir = ConfigurationManager.AppSettings["TargetDir"];
                sDatabaseLocation = ConfigurationManager.AppSettings["DatabaseLocation"];
                sFtpSite= ConfigurationManager.AppSettings["FtpSite"];
	            sFtpSiteUserName= ConfigurationManager.AppSettings["FtpSiteUserName"];
	            sFtpSitePassword= ConfigurationManager.AppSettings["FtpSitePassword"];
	            sHostKey= ConfigurationManager.AppSettings["HostKey"];
	            sDebug= ConfigurationManager.AppSettings["Debug"];
                sPostToSftp = ConfigurationManager.AppSettings["PostToSftp"];

                if (sDebug == "true") { bDebug = true; }
                 if (sPostToSftp == "true") { bPostToSftp = true; }


                //msLogFile = ConfigurationManager.AppSettings["logfile"];

                if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent + "SourceDir = " + sSourceDir); }
                if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent + "TargetDir = " + sTargetDir); }
                if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent + "DatabaseLocation = " + sDatabaseLocation); }
                if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent + "PostToSftp = " + bPostToSftp); }

                // Parse source folders
                String[] sSourcefolders = Directory.GetDirectories(sSourceDir);
                foreach (String sSourcefolder in sSourcefolders)
                {
                    //String lastFolderName = Path.GetFileName(Path.GetDirectoryName(folder)); // = live properties
                   // String lastFolderName = Path.GetDirectoryName(folder);
                   String sLastSourceFolderName = sSourcefolder.Split(Path.DirectorySeparatorChar).Last();

                    //Console.WriteLine("lastFolderName  = " + sLastSourceFolderName);
                    // Check if folder exists in target location, and if not create it
                    sNewFolder =sTargetDir+"\\"+sLastSourceFolderName;
                    if (!Directory.Exists(sNewFolder)) {
                        clsLog_.mLog(Constants.gcInfo, "Creating new folder " + sNewFolder);
                        Directory.CreateDirectory(sNewFolder);
                    }
                    // Process images
                     bProcessImage(sSourcefolder + "\\Main","0.jpg", sNewFolder);
                     bProcessImage(sSourcefolder + "\\Medium","1.jpg", sNewFolder);
                     bProcessImage(sSourcefolder + "\\Medium","2.jpg", sNewFolder);
                     bProcessImage(sSourcefolder + "\\Medium","3.jpg", sNewFolder);

                }

                // Refresh data
                bOk = bRefreshData(sDatabaseLocation, bDebug);
                // Export to xml
                if (bOk){ bOk = bExportDataToXML(sDatabaseLocation, bDebug);}
                // Post of sftp site
                if (bOk && bPostToSftp) { bOk = bsFtp(sFtpSite, sFtpSiteUserName, sFtpSitePassword, sHostKey, sDatabaseLocation, false, bDebug, false); }
            }
            catch (Exception e)
            {
                clsError_.mLogError("Problem running onelan interface", "onelan", "Main", e, msVersionData, msLogFile, bAlertUser);

            }
            finally
            {
                clsLog_.mLog(Constants.gcInfo, "Processing complete, status = " + bOk); 
                Console.WriteLine("Processing complete, status = " + bOk);
            }

        // End of main process
        }


        //
        // Process a single image
        //
        public static Boolean bProcessImage(String sSourceFolder, String sSourceImage, String sNewFolder)
        {

            clsLog clsLog_ = new clsLog(msLogFile, msVersionData);
            clsError clsError_ = new clsError();

            // Flag to indicate file should be copied    
            Boolean bDoCopy = false;
            try
            {
                // Test
              //  if (sSourceFolder=="C:\\Projects\\mantis\\1541OneLan\\LiveProperties\\12901\\Main") {
               //     Console.WriteLine("onelane.bProcessImage, sSourceFolder = " + sSourceFolder);
                  //  Console.WriteLine("onelane.bProcessImage,  sSourceImage= "+sSourceImage);
                 //   Console.WriteLine("onelane.bProcessImage,  sNewFolder= " + sNewFolder);
               
                    // Specify source and target file name and location
                    // Note  0.jpg => 0.jpg_s.jpg
                    String sSourceImageLocation = sSourceFolder + "\\" + sSourceImage;
                    String sTargetImageLocation = sNewFolder+"\\"+sSourceImage+"_s.jpg";

                    //If source file does not exist then quit
                    if (!File.Exists(sSourceImageLocation)){ return true;}
                    // Check if file is not present in target location
                     if (!File.Exists(sTargetImageLocation))
                     { 
                         bDoCopy= true;
                     }
                    else
                     {
                    // If file exists then need to compare date stamps
                         DateTime DateTimeSource = File.GetLastWriteTime(sSourceImageLocation);
                         DateTime DateTimeTarget = File.GetLastWriteTime(sTargetImageLocation);
                         int result = DateTime.Compare(DateTimeSource, DateTimeTarget);
                         if (result > 0) { bDoCopy = true; }
                     }

                    // if bDoCopy FlagsAttribute has been set ToolBar true then copy the file over
                    if( bDoCopy) {
                        clsLog_.mLog(Constants.gcInfo, "onelan.bProcessImage,  Copying file"+sSourceImageLocation);
                        File.Copy(sSourceImageLocation,sTargetImageLocation,true);
                 //   }

                }

                return true;
            }
            catch (Exception e)
            {

                clsError_.mLogError("Problem processing image", "onelan", "Main", e, msVersionData, msLogFile, bAlertUser);
                return false;
            }

        }
        //
        // Refresh data
        //
        public static Boolean bRefreshData(String sDatabaseLocation, Boolean bDebug)
        {

            clsLog clsLog_ = new clsLog(msLogFile, msVersionData);
            clsError clsError_ = new clsError();

            String sSql = null;
            ADODB.Connection conn = new ADODB.Connection();
            object RecordsAffected = null;
     
            try
            {

                clsLog_.mLog(Constants.gcInfo, "Refreshing data ..."); 

                if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent + "onelan.bRefreshData, connection =" + sDatabaseLocation + gsOneLanDatabase); }
                // Open connection to database
                conn.Open(gcDSN + sDatabaseLocation + "\\" + gsOneLanDatabase);


                 //1.1 Purge accom table for properities no longer current
                sSql = "DELETE FROM accom a" +
                       " WHERE NOT EXISTS (" +
                       "  SELECT 1" +
                       "    FROM PropertyDetailsRC rc" +
                       "   WHERE rc.propid = a.propid" +
                       "     AND rc.sold =0 and pending = 0)";

                conn.Execute(sSql, out RecordsAffected, -1);

                                //1.2 Purge Property table for properities no longer current
                sSql = "DELETE FROM Property p" +
                        " WHERE NOT EXISTS (" +
                        "  SELECT 1" +
                        "    FROM PropertyDetailsRC rc" +
                        "   WHERE rc.propid = p.propid" +
                        "     AND rc.sold =0 AND pending =0)";
	            conn.Execute(sSql, out RecordsAffected, -1);
	  
                //1.3 Add new properties	  
                sSql = "INSERT INTO Property (" +
                        " PropID" +
                        ",PropertyAddress1" +
                        ",PropertyAddress2" +
                        ",PropertyAddress3" +
                        ",PropertyAddress4" +
                        ",PostCode" +
                        ",Price" +
                        ",OffersOverEtc" +
                        ",RCCWOffice" +
                        ",UnderOffer" +
                        ",ClosingDate )";
     
                sSql = sSql  + "SELECT " +
                        " rc.PropID" +
                        ",rc.PropertyAddress1" +
                        ",rc.PropertyAddress2" +
                        ",rc.PropertyAddress3" +
                        ",rc.PropertyAddress4" +
                        ",rc.PostCode" +
                        ",rc.Price" +
                        ",rc.OffersOverEtc" +
                        ",rc.RCCWOffice" +
                        ",rc.UnderOffer" +
                        ",rc.ClosingDate" +
                        "  FROM PropertyDetailsRC rc" +
                        " WHERE rc.sold=0" +
                        "   AND pending =0" +
                        " AND NOT EXISTS (" +
                        "  SELECT 1" +
                        "    FROM Property p2" +
                        "   WHERE p2.propid = rc.propid)";
	            conn.Execute(sSql, out RecordsAffected, -1);	  
	  
                //1.4 Update property details to ensure current with main system	  
                sSql = "UPDATE property p," +
                        "       PropertyDetailsRC rc  SET " +
                        "       p.PropertyAddress1 = rc.PropertyAddress1" +
                        "      ,p.PropertyAddress2 = rc.PropertyAddress2" +
                        "      ,p.PropertyAddress3 = rc.PropertyAddress3" +
                        "      ,p.PropertyAddress4 = rc.PropertyAddress4" +
                        "      ,p.PostCode= rc.PostCode" +
                        "      ,p.Price= rc.Price" +
                        "      ,p.OffersOverEtc= rc.OffersOverEtc" +
                        "      ,p.ClosingDate= rc.ClosingDate" +
                        " WHERE p.Propid = rc.Propid" +
                        "   AND (" +
                        "       p.PropertyAddress1 <> rc.PropertyAddress1" +
                        "    OR p.PropertyAddress2 <> rc.PropertyAddress2" +
                        "    OR p.PropertyAddress3 <> rc.PropertyAddress3" +
                        "    OR p.PropertyAddress4 <> rc.PropertyAddress4" +
                        "    OR p.PostCode <> rc.PostCode" +
                        "    OR p.Price <> rc.Price" +
                        "    OR p.OffersOverEtc <> rc.OffersOverEtc" +
                        "    OR p.ClosingDate <> rc.ClosingDate" +
                        " )";

                // Fix the closing date
                //1.4.1 Update null closing date to 00:00:00

                sSql = "Update Property" +
                " Set closingdate = '00:00:00'" +
                " Where closingdate Is Null or format(closingDate,'dd-mmm-yyyy')=''";
  	            conn.Execute(sSql, out RecordsAffected, -1);
  
  
                //1.4 Add new ACC 
                sSql = "  INSERT INTO ACCOM (Propid,Accom, Website)" +
                        " SELECT rca.Propid,rca.Accom, rca.Website" +
                        "   FROM ACCOMRC rca," +
                        "        property p " +
                        " WHERE p.propid= rca.propid" +
                        " AND NOT EXISTS (" +
                        "     SELECT 1" +
                        "       FROM accom a" +
                        "      WHERE a.propid = p.propid)";
  	            conn.Execute(sSql, out RecordsAffected, -1);

                conn.Close();


                return true;
            }
            catch (Exception e)
            {
                clsError_.mLogError("Problem running onelan interface", "onelan", "bRefreshData", e, msVersionData, msLogFile, true);
                return false;
            }

        }

        // Output data to xml
        public static Boolean bExportDataToXML(String sDatabaseLocation, Boolean bDebug)
        {

            clsLog clsLog_ = new clsLog(msLogFile, msVersionData);
            clsError clsError_ = new clsError();

            ADODB.Connection conn = new ADODB.Connection();
            ADODB.Recordset rs = new ADODB.Recordset();

            clsLog_.mLog(Constants.gcInfo, "Exporting data to xml ...");

           
            String sOneLanPropertyfile = sDatabaseLocation + "\\xml\\OneLanProperty.xml";
            String sOneLanAccomfile = sDatabaseLocation + "\\xml\\OneLanAccom.xml";

            if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent + "onelan.bExportDataToXML, sOneLanPropertyfile = " + sOneLanPropertyfile); }
            if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent + "onelan.bExportDataToXML, sOneLanAccomfile = " + sOneLanAccomfile); }

            try
            {
                // Open connection to database
                conn.Open(gcDSN + sDatabaseLocation + "\\" + gsOneLanDatabase);

                // Remove previous versions of the files
                if (File.Exists(sOneLanPropertyfile)) {File.Delete(sOneLanPropertyfile);}
                if (File.Exists(sOneLanAccomfile)) {File.Delete(sOneLanAccomfile);}

                // Open the recordset
                rs.Open("SELECT * FROM Property", conn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic, -1);
                rs.Save(sOneLanPropertyfile, ADODB.PersistFormatEnum.adPersistXML);
                rs.Close();

                rs.Open("SELECT * FROM Accom", conn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic, -1);
                rs.Save(sOneLanAccomfile, ADODB.PersistFormatEnum.adPersistXML);
                rs.Close();

                conn.Close();

                return true;
            }
            catch (Exception e)
            {
               clsError_.mLogError("Problem exporting data to xml", "onelan", "bExportDataToXML", e, msVersionData, msLogFile, true);
               return false;
            }

        }
        // End of bExportDataToXML


        // Post data via sftp
        //
        public static Boolean bsFtp(String sFtpSite, String sUserName, String sPassword, String sSshHostKeyFingerprint, String sDatabaseLocation, Boolean bSearchAll = false, Boolean bDebug = false, Boolean bExposeCredentials = false)
        {

            clsLog clsLog_ = new clsLog(msLogFile, msVersionData);
            clsError clsError_ = new clsError();

           // String sAuditRootUCase = "";
           // String sTRAMSFtpAuditRootDir = "";
            // String sTRAMSFtpRootDir = "";
            SynchronizationResult synchronizationResult;


            try
            {

                clsLog_.mLog(Constants.gcInfo, "Posting data to sftp site ...");

                //clsLog_.mLog(Constants.gcInfo, "");
               // clsLog_.mLog(Constants.gcInfo, "onelan.bsFtp, sftp (winscp) process starting **************");
                if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent +"onelan.bsFtp, sftpSite = " + sFtpSite); }
                if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent + "onelan.bsFtp, sUserName = " + sUserName); }
                if (bExposeCredentials)
                {
                    clsLog_.mLog(Constants.gcInfo, sIndent +"************ Credentials *****************");
                    clsLog_.mLog(Constants.gcInfo, sIndent +"onelan.bsFtp, sPassword = " + sPassword);
                    clsLog_.mLog(Constants.gcInfo, sIndent +"onelan.bsFtp, sSshHostKeyFingerprint = " + sSshHostKeyFingerprint);
                    clsLog_.mLog(Constants.gcInfo, "");
                }
                if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent + "onelan.bsFtp, local root folder = " + sDatabaseLocation); }
               // clsLog_.mLog(Constants.gcInfo, "onelan.bsFtp, SearchAll = " + bSearchAll);
               // clsLog_.mLog(Constants.gcInfo, "onelan.bsFtp, Debug = " + bDebug);


                // Setup session options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Sftp,
                    HostName = sFtpSite,
                    UserName = sUserName,
                    Password = sPassword,
                    SshHostKeyFingerprint = sSshHostKeyFingerprint

                };

                using (Session session = new Session())
                {
                    // Connect
                    if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent + "onelan.bsFtp, connecting ..."); }
                    session.Open(sessionOptions);

                    // Upload files
                    if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent + "onelan.bsFtp, uploading ..."); }

                    // Get list of directories to transfer

                    // Will continuously report progress of synchronization
                    //if (bDebug) { session.FileTransferred += FileTransferred; }

                    if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent + "onelan.bsFtp, Posting images ..."); }

                    synchronizationResult = session.SynchronizeDirectories(SynchronizationMode.Remote, sDatabaseLocation + "\\images", "images", false);

                    if (bDebug) { clsLog_.mLog(Constants.gcInfo, sIndent + "onelan.bsFtp, Processing xml files ..."); }
                    // Synchronise the xml directory
                    synchronizationResult = session.SynchronizeDirectories(SynchronizationMode.Remote, sDatabaseLocation + "\\xml", "xml", false);

                    // Clean up
                    session.Dispose();
                }



               // clsLog_.mLog(Constants.gcInfo, "********** Completed sftp processing **********");
                return true;
            }
            catch (Exception e)
            {
                clsError_.mLogError("Problem running sftpTransfer", "onelan", "bsFtp", e, msVersionData, msLogFile, true);
                return false;
            }

            finally
            {


            }


        }




    }
    // End of class
}
// End of NameSpace
