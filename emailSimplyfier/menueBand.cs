using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace emailSimplyfier
{
    public partial class menueBand
    {
        //Used in determing which language to display
        static private string systemLanguage = System.Globalization.CultureInfo.CurrentUICulture.Name;

        static private string emailPath = "";
        static private string attachmentPath = "";

        //The eventhandler of the first button ("E-Mail archivieren") in the gallery list
        private void ArchiveEmail_Click(object sender, RibbonControlEventArgs e)
        {
            //Specify the current window the add-in is used in
            Microsoft.Office.Interop.Outlook.Inspector m = e.Control.Context as Inspector;

            //Declare an object of the mail itself, which is displayed in the current window
            Microsoft.Office.Interop.Outlook.MailItem currentMail = m.CurrentItem as MailItem;

            //Sets the emailPath
            CheckAndSetFolderPath(1);


            ////Attributes, which can be used in file name when saving the .msg file
            //--------------------------------------------------------------------------------
            ////Sender e-mail address
            //string senderAddress = currentMail.SenderEmailAddress;
            //
            ////Number of attachments
            //string numberOfAttachments = currentMail.Attachments.Count.ToString();
            //--------------------------------------------------------------------------------


            //These 3 variables should be enough for the .msg filename to be unique
            //Name of the sender
            string senderName = currentMail.SenderName;
            //Subject of the email
            string subjectLine = currentMail.Subject;
            //Receive date
            string receivedDate = currentMail.ReceivedTime.ToString("dd MM yyyy HH mm ss");
            
            
            //The actual saving of the file
            try
            {
                //Concatination of the folderpath + sequence of name variables + .msg to form the full path, where you want to save the file
                string folderPath = emailPath + senderName + " , " + subjectLine + " , " + receivedDate +".msg";


                //Check, if the file already exists or not and proceed depending on the outcome of the check
                if (File.Exists(folderPath) == false)
                {
                    currentMail.SaveAs(folderPath, OlSaveAsType.olMSG);

                    //Get the first two characters of the language ( ex. "de-DE" / "de-AT" / "de-CH" can understand language "de", language check is easier that way)
                    switch (systemLanguage.Substring(0,2))
                    {   
                        //More languages can be added in the future
                        case "de":
                            MessageBox.Show("Archivieren erfolgreich.");
                            break;
                        case "en":
                            MessageBox.Show("Archiving successful.");
                            break;
                    }
                    
                }
                else
                {
                    switch (systemLanguage.Substring(0, 2))
                    {
                        case "de":
                            MessageBox.Show("Archivieren fehlgeschlagen. \n \n" 
                                
                                + "Es existiert bereits eine Datei mit diesem Namen.");
                            break;
                        case "en":
                            MessageBox.Show("Archive failed. \n \n" 
                                
                                + "A file with this name already exists.");
                            break;
                    }
                    
                }
            }
            //This catch covers the case, if you want to save a file into a path which does not exist
            catch (DirectoryNotFoundException)
            {
                switch (systemLanguage.Substring(0, 2))
                {
                    case "de":
                        MessageBox.Show("Archivieren fehlgeschlagen. \n \n" 

                            + "Der angegebene Ordner wurde nicht gefunden. \n" +
                            "Bitte ändern Sie den Pfad oder stellen Sie sicher, dass die folgende Struktur exisistiert: \n \n"

                            + emailPath);
                        break;
                    case "en":
                        MessageBox.Show("Archive failed. \n \n" 

                            + "The specified folder was not found. \n" +
                            " Please change the path or make sure, that the following structure is existing: \n \n"
                            
                            + emailPath);
                        break;
                }
                
            }
    }

        //The eventhandler of the second button (PDF verarbeiten) in the gallery list
        private void ProcessPDF_Click(object sender, RibbonControlEventArgs e)
        {   
            //Specify the current window the add-in is used in
            Microsoft.Office.Interop.Outlook.Inspector m = e.Control.Context as Inspector;

            //Declare an object of the mail, which is displayed in the current window
            Microsoft.Office.Interop.Outlook.MailItem currentMail = m.CurrentItem as MailItem;


            //Variables needed in the following 
            string failedFileNames = "";
            string successFileNames = "";
            ArrayList listOfSuccessfullFiles = new ArrayList();
            ArrayList listOfFailedFiles = new ArrayList();
            
            
            //Try Catch to check wether the pathfolder exists or not
            try
            {
                //This check is needed if you put the button in Microsoft.Outlook.Mail.Explorer instead of Microsoft.Outlook.Mail.Read 
                if (currentMail != null)
                {   
                    //Check if email has attachments
                    
                    //The email has atleast one attachment
                    if (currentMail.Attachments.Count > 0)
                    {
                        //Sets the attachmentPath
                        CheckAndSetFolderPath(2);


                        //Loop over Attachmentlist and decide what to do with the attachment(s)

                        //Possible outcomes:
                        //1. There is no file with the same name as the current Attachment in scope 
                        //-> save the attachment as a file and add the attachment name to listOfSuccessfullFiles for later use

                        //2. There is a file [..]                                                   
                        //-> dont save the attachment as a file but add the attachment name listOfFailedFiles

                        foreach (Attachment item in currentMail.Attachments)
                        {   
                            //1.
                            if (File.Exists(attachmentPath+item.FileName) == false)
                            {
                                listOfSuccessfullFiles.Add(item.FileName);
                                item.SaveAsFile(attachmentPath + item.FileName);
                            }
                            //2.
                            else
                            {
                                listOfFailedFiles.Add(item.FileName);
                            }                           
                        }

                        //Now decide what the MessageBox text is, based on the contents of listOfSuccessfullFiles & listOfFailedFiles

                        //A file was / All files were saved successfully
                        if (listOfFailedFiles.Count == 0)
                        {
                            //One file was saved successfully
                            if (currentMail.Attachments.Count == 1)
                            {   
                                switch (systemLanguage.Substring(0, 2))
                                {
                                    case "de":
                                        MessageBox.Show("Eine Datei erfolgreich heruntergeladen: \n"
                                            + listOfSuccessfullFiles[0]);
                                        break;
                                    case "en":
                                        MessageBox.Show("One File was successfully downloaded: : \n"
                                            + listOfSuccessfullFiles[0]);
                                        break;
                                }

                            }
                            //Multiple files were saved successfully
                            else
                            {
                                successFileNames = ConcationationOfFileNames(listOfSuccessfullFiles);
                                switch (systemLanguage.Substring(0, 2))
                                {
                                    case "de":
                                        MessageBox.Show(currentMail.Attachments.Count + " Dateien erfolgreich heruntergeladen: \n"
                                            +successFileNames);
                                        break;
                                    case "en":
                                        MessageBox.Show(currentMail.Attachments.Count + " Files successfully downloaded: \n"
                                            + successFileNames);
                                        break;
                                }

                            }
                        }
                        //Not every file was saved successfully
                        else
                        {
                            int successfullyDownloaded = currentMail.Attachments.Count - listOfFailedFiles.Count;

                            //No attachment(s) could be saved
                            if (successfullyDownloaded == 0)
                            {   
                                //Email has only one attachment which couldn't be saved
                                if(listOfFailedFiles.Count == 1)
                                {
                                    switch (systemLanguage.Substring(0, 2))
                                    {
                                        
                                        case "de":
                                            MessageBox.Show("Download der Datei fehlgeschlagen. \n \n" 

                                                + "Eine Datei mit entsprechendem Namen existiert bereits: \n" 
                                                + listOfFailedFiles[0]);
                                            break;
                                        case "en":
                                            MessageBox.Show("File download failed. \n \n" 

                                                + "A file with the same name already exists: \n" 
                                                + listOfFailedFiles[0]);
                                            break;
                                    }
                                }
                                //Email has multiple attachments and not a single one could be saved
                                else
                                {
                                    failedFileNames = ConcationationOfFileNames(listOfFailedFiles);
                                    switch (systemLanguage.Substring(0, 2))
                                    {
                                        case "de":
                                            MessageBox.Show("Download der Dateien fehlgeschlagen. \n \n"  
                                                
                                                + listOfFailedFiles.Count + " Dateien mit entsprechenden Namen existieren bereits: \n" 
                                                + failedFileNames);
                                            break;
                                        case "en":
                                            MessageBox.Show("Download of the files failed. \n \n" 

                                                + listOfFailedFiles.Count + " Files with corresponding names already exist: \n" 
                                                + failedFileNames);
                                            break;
                                    }
                                }
                                
                            }
                            //Out of all attachments one/multiple were saved and one/multiple couldnt be saved
                            else
                            {
                                //One attachment out of multiple was saved
                                if(successfullyDownloaded == 1)
                                {
                                    //One attachment couldn't be saved (number of attachments = 2 -> 1 saved 1 not saved)
                                    if(listOfFailedFiles.Count == 1)
                                    {

                                        switch (systemLanguage.Substring(0, 2))
                                        {
                                            case "de":
                                                MessageBox.Show("Dateien teilweise heruntergeladen. \n \n" 

                                                    + "Es wurde eine Datei erfolgreich heruntergeladen: \n" 
                                                    + listOfSuccessfullFiles[0] + "\n \n"

                                                    + "Eine Datei mit entsprechendem Namen existiert bereits: \n" 
                                                    + listOfFailedFiles[0]);
                                                break;
                                            case "en":
                                                MessageBox.Show("Files partially downloaded. \n \n" 

                                                    + "A File was successfully downloaded: \n" 
                                                    + listOfSuccessfullFiles[0] + "\n \n"

                                                    + "A File with this name already existed: \n" 
                                                    + listOfFailedFiles[0]);
                                                break;
                                        }
                                    }
                                    //Multiple attachments couldn't be saved (number of attachments >=3 -> 1 saved >=2 not saved)
                                    else
                                    {
                                        failedFileNames = ConcationationOfFileNames(listOfFailedFiles);
                                        switch (systemLanguage.Substring(0, 2))
                                        {
                                            case "de":
                                                MessageBox.Show("Dateien teilweise heruntergeladen. \n \n" 

                                                    + "Es wurde eine Datei erfolgreich heruntergeladen: \n" 
                                                    + listOfSuccessfullFiles[0] + "\n \n"

                                                    + listOfFailedFiles.Count + " Dateien mit entsprechenden Namen existierten bereits: \n" 
                                                    + failedFileNames);
                                                break;
                                            case "en":
                                                MessageBox.Show("Files partially downloaded. \n \n"

                                                    + "A File was successfully downloaded: \n"  
                                                    + listOfSuccessfullFiles[0] + "\n \n"

                                                    + listOfFailedFiles.Count + " Files with corresponding names already existed: \n" 
                                                    + failedFileNames);
                                                break;
                                        }
                                    }
                                    
                                }
                                //Multiple attachments but not all were saved
                                else
                                {
                                    //One attachment couldn't be saved (number of attachments >=3 -> >=2 saved 1 not saved)
                                    if (listOfFailedFiles.Count == 1)
                                    {
                                        successFileNames = ConcationationOfFileNames(listOfSuccessfullFiles);
                                        switch (systemLanguage.Substring(0, 2))
                                        {
                                            case "de":
                                                MessageBox.Show("Dateien teilweise heruntergeladen. \n \n" 

                                                    + "Es wurden " + successfullyDownloaded + " Dateien erfolgreich heruntergeladen: \n" 
                                                    + successFileNames + "\n"

                                                    + "Eine Datei mit entsprechendem Namen existiert bereits: \n" 
                                                    + listOfFailedFiles[0]);
                                                break;
                                            case "en":
                                                MessageBox.Show("Files partially downloaded. \n \n" 

                                                    + successfullyDownloaded + "Files were successfully downloaded. \n \n"

                                                    + "A File with this name already existed: \n" 
                                                    + listOfFailedFiles[0]);
                                                break;
                                        }
                                    }
                                    //Multiple attachments couldn't be saved (number of attachments >=4 -> >=2 saved >=2 not saved)
                                    else
                                    {
                                        failedFileNames = ConcationationOfFileNames(listOfFailedFiles);
                                        successFileNames = ConcationationOfFileNames(listOfSuccessfullFiles);
                                        switch (systemLanguage.Substring(0, 2))
                                        {
                                            case "de":
                                                MessageBox.Show("Dateien teilweise heruntergeladen. \n \n" 

                                                    + "Es wurden " + successfullyDownloaded + " Dateien erfolgreich heruntergeladen: \n"
                                                    + successFileNames + "\n"

                                                    + listOfFailedFiles.Count + " Dateien mit entsprechenden Namen existierten bereits: \n" 
                                                    + failedFileNames);
                                                break;
                                            case "en":
                                                MessageBox.Show("Files partially downloaded. \n \n" 

                                                    + successfullyDownloaded + "Files were successfully downloaded: \n"
                                                    + successFileNames + "\n"

                                                    + listOfFailedFiles.Count + " Files with corresponding names already existed: \n" 
                                                    + failedFileNames);
                                                break;
                                        }
                                    }
                                }
                                
                            }
                        }
                        
                    }
                    //The email has no attachments
                    else
                    {   
                        switch (systemLanguage.Substring(0, 2))
                        {
                            case "de":
                                MessageBox.Show("Diese E-Mail hat keine Anhänge.");
                                break;
                            case "en":
                                MessageBox.Show("This e-mail has no attachments.");
                                break;
                        }
                    }
                }
            }
            //Proceed to inform the user if the folder is not existing
            catch (DirectoryNotFoundException)
            {
                switch (systemLanguage.Substring(0, 2))
                {
                    case "de":
                        MessageBox.Show("Archivieren fehlgeschlagen. \n \n"

                            + "Der angegebene Ordner wurde nicht gefunden. \n" +
                            "Bitte ändern Sie den Pfad oder stellen Sie sicher, dass die folgende Struktur exisistiert: \n \n"

                            + attachmentPath);
                        break;
                    case "en":
                        MessageBox.Show("Archive failed. \n \n"

                            + "The specified folder was not found. \n" +
                            " Please change the path or make sure, that the following structure is existing: \n \n"

                            + attachmentPath);
                        break;
                }
            }
        }



        //Concatinate given filenames to print them easier on the messagebox
        private string ConcationationOfFileNames(ArrayList list)
        {
            string finalString = "";

            foreach (string fileName in list)
            {
                finalString = finalString + fileName + "\n";
            }

            return finalString;
        }

        private void CheckAndSetFolderPath(int caseDecider)
        {
            //Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)))
            string pathTemp = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\emailSimplyfier\";

            string currentUserName = Environment.UserName;

            switch (caseDecider)
            {
                //ArchiveEmail_Click was clicked
                case 1:
                    
                    //Check if emailSimplyfier exists, if not create
                    if(Directory.Exists(pathTemp) == true)
                    {
                        //Check if ArchiveEmail_Click subfolder exists, if not create
                        pathTemp = pathTemp + @"\Archivierte E-Mails\";
                        if(Directory.Exists(pathTemp) == true)
                        {
                            emailPath = pathTemp;
                        }
                        else
                        {
                            MessageBox.Show("Ordner Archivierte E-Mails wurde unter " + currentUserName + @"\Dokumente\emailSimplyfier " + "erstellt.");
                            Directory.CreateDirectory(pathTemp);
                            emailPath = pathTemp;
                        }
                    }
                    else
                    {
                        Directory.CreateDirectory(pathTemp);
                        //Check if ArchiveEmail_Click subfolder exists, if not create
                        pathTemp = pathTemp + @"\Archivierte E-Mails\";
                        if (Directory.Exists(pathTemp) == true)
                        {
                            emailPath = pathTemp;
                        }
                        else
                        {
                            MessageBox.Show("Ordner "+ currentUserName +  @"\Dokumente\emailSimplyfier\Archivierte E-Mails\ " +  "wurden erstellt.");
                            Directory.CreateDirectory(pathTemp);
                            emailPath = pathTemp;
                        }
                    }
                    break;
                case 2:
                    //Check if emailSimplyfier exists, if not create
                    if (Directory.Exists(pathTemp) == true)
                    {
                        //Check if ProcessPDF_Click subfolder exists, if not create
                        pathTemp = pathTemp + @"\Heruntergeladene Anhänge\";
                        if (Directory.Exists(pathTemp) == true)
                        {
                            attachmentPath = pathTemp;
                        }
                        else
                        {
                            MessageBox.Show("Ordner Heruntergeladene Anhänge wurde unter " + currentUserName + @"\Dokumente\emailSimplyfier " + "erstellt.");
                            Directory.CreateDirectory(pathTemp);
                            attachmentPath = pathTemp;
                        }
                    }
                    else
                    {
                        Directory.CreateDirectory(pathTemp);

                        //Check if ArchiveEmail_Click subfolder exists, if not create
                        pathTemp = pathTemp + @"\Heruntergeladene Anhänge\";
                        if (Directory.Exists(pathTemp) == true)
                        {
                            emailPath = pathTemp;
                        }
                        else
                        {   
                            MessageBox.Show("Ordner " + currentUserName + @"\Dokumente\emailSimplyfier\Heruntergeladene Anhänge\ " + "wurden erstellt.");
                            Directory.CreateDirectory(pathTemp);
                            emailPath = pathTemp;
                        }
                    }

                    break;
            }
        }


        ////this is my testbutton to test certain functions faster, can be deleted in final version
        //private void button3_Click(object sender, RibbonControlEventArgs e)
        //{
        //    //Specify the current window the add-in is used in
        //    Microsoft.Office.Interop.Outlook.Inspector m = e.Control.Context as Inspector;

        //    //Declare an object of the mail, which is displayed in the current window
        //    Microsoft.Office.Interop.Outlook.MailItem currentMail = m.CurrentItem as MailItem;

        //    MessageBox.Show(currentMail.Body);
            
        //}



        //leave empty, because the events happen in the gallery items (archiveEmail, processPDF etc.) not in the gallery itself
        //got autogenerated, when creating the gallery. might produce bugs, if you delete it, so i left it

        private void emSimGallery_Click(object sender, RibbonControlEventArgs e)
        {
        }
        
        private void menueBand_Load(object sender, RibbonUIEventArgs e)
        {
        }
    }
}
