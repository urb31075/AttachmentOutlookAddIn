// --------------------------------------------------------------------------------------------------------------------
// <copyright file="AttachmentUserControl.cs" company="urb31075">
// All Right Reserved  
// </copyright>
// <summary>
//   Defines the AttachmentUserControl type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------
namespace AttachmentOutlookAddIn
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// The attachment user control.
    /// </summary>
    public partial class AttachmentUserControl : UserControl
    {
        /// <summary>
        /// The save path.
        /// </summary>
        private const string SavePath = @"C:\EmailAttachments";

        /// <summary>
        /// The error list.
        /// </summary>
        private List<string> errorList;

        /// <summary>
        /// Initializes a new instance of the <see cref="AttachmentUserControl"/> class.
        /// </summary>
        public AttachmentUserControl()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// The attachment user control_ load.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private void AttachmentUserControlLoad(object sender, EventArgs e)
        {
            this.FolderComboBox.Items.Clear();
            var count = 0;
            var selectedIndex = 0;
            foreach (dynamic folder in ThisAddIn.thisApplication.GetNamespace("MAPI").Folders)
            {
                var subFolders = this.GetFolder(folder.FolderPath);
                foreach (Outlook.MAPIFolder subFolder in subFolders.Folders)
                {
                    this.FolderComboBox.Items.Add(subFolder);
                    //// if (subFolder.FullFolderPath.Contains("r.ugryumov@GASP.RU") && subFolder.FullFolderPath.Contains("Входящие"))
                    if (subFolder.FullFolderPath.Contains("Входящие"))
                    {
                        selectedIndex = count;
                    }

                    count++;
                }

                Marshal.ReleaseComObject(subFolders);
            }

            this.FolderComboBox.SelectedIndex = selectedIndex;
        }

        /// <summary>
        /// The get folder.
        /// </summary>
        /// <param name="folderPath">
        /// The folder path.
        /// </param>
        /// <returns>
        /// Получение списка папок оутлоока
        /// </returns>
        private Outlook.Folder GetFolder(string folderPath)
        {
            try
            {
                folderPath = folderPath.TrimStart("\\".ToCharArray()); // Remove leading "\" characters.
                var folders = folderPath.Split("\\".ToCharArray()); // Split the folder path into individual folder names.
                var returnFolder = ThisAddIn.thisApplication.Session.Folders[folders[0]] as Outlook.Folder;
                if (returnFolder != null)
                {
                    for (int i = 1; i < folders.Length; i++)
                    {
                        var folderName = folders[i];
                        if (returnFolder == null)
                        {
                            continue;
                        }

                        var subFolders = returnFolder.Folders;
                        returnFolder = subFolders[folderName] as Outlook.Folder;
                    }
                }

                return returnFolder;
            }
            catch (Exception ex)
            {
                if (this.errorList == null)
                {
                    this.errorList = new List<string>();
                }

                this.errorList.Add(MethodBase.GetCurrentMethod().Name + " Error message : " + ex.Message);
                this.errorList.Add(MethodBase.GetCurrentMethod().Name + " StackTrace " + ex.StackTrace);
                return null;
            }
        }

        /// <summary>
        /// The extract button_ click.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private void ExtractButtonClick(object sender, EventArgs e)
        {
            var currentCursor = Cursor.Current;
            try
            {
                this.Cursor = Cursors.WaitCursor;
                Application.DoEvents();
                if (this.FolderComboBox.SelectedItem != null)
                {
                    var inboxFolder = (Outlook.MAPIFolder)this.FolderComboBox.SelectedItem;
                    ThisAddIn.thisApplication.ActiveExplorer().CurrentFolder = inboxFolder;
                    ThisAddIn.thisApplication.ActiveExplorer().CurrentFolder.Display();

                    var attachmentDataList = this.ExtractAttachment(inboxFolder);
                    this.InfoListBox.Items.Clear();
                    foreach (var attachment in attachmentDataList)
                    {
                        this.InfoListBox.Items.Add(attachment);    
                    }

                    this.InfoToolStripStatusLabel.Text = string.Format("Обнаружено: {0}", attachmentDataList.Count);
                    MessageBox.Show(@"Вложения скопированы в " + SavePath, @"Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                this.InfoListBox.Items.Add(MethodBase.GetCurrentMethod().Name + " Error message : " + ex.Message);
                this.InfoListBox.Items.Add(MethodBase.GetCurrentMethod().Name + " StackTrace " + ex.StackTrace);
            }
            finally
            {
                this.Cursor = currentCursor;
                Application.DoEvents();
            }
        }

        /// <summary>
        /// The clear button_ click.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private void ClearButtonClick(object sender, EventArgs e)
        {
            var currentCursor = Cursor.Current;
            try
            {
                this.Cursor = Cursors.WaitCursor;
                Application.DoEvents();

                if (this.FolderComboBox.SelectedItem != null)
                {
                    var inboxFolder = (Outlook.MAPIFolder)this.FolderComboBox.SelectedItem;
                    var msg = string.Format("Будут удалены вложения из '{0}'. Продолжить?", inboxFolder.FolderPath);
                    if (MessageBox.Show(msg, @"Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        return;
                    }

                    ThisAddIn.thisApplication.ActiveExplorer().CurrentFolder = inboxFolder;
                    ThisAddIn.thisApplication.ActiveExplorer().CurrentFolder.Display();

                    var attachmentDataList = this.ClearAttachment(inboxFolder);
                    this.InfoListBox.Items.Clear();
                    foreach (var attachment in attachmentDataList)
                    {
                        this.InfoListBox.Items.Add(attachment);
                    }

                    this.InfoToolStripStatusLabel.Text = string.Format("Обнаружено: {0}", attachmentDataList.Count);
                }
            }
            catch (Exception ex)
            {
                this.InfoListBox.Items.Add(MethodBase.GetCurrentMethod().Name + " Error message : " + ex.Message);
                this.InfoListBox.Items.Add(MethodBase.GetCurrentMethod().Name + " StackTrace " + ex.StackTrace);
            }
            finally
            {
                this.Cursor = currentCursor;
                Application.DoEvents();
            }
        }

        /// <summary>
        /// The get recived button click.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private void GetReceivedButtonClick(object sender, EventArgs e)
        {
            var currentCursor = Cursor.Current;
            try
            {
                this.Cursor = Cursors.WaitCursor;
                Application.DoEvents();
                if (this.FolderComboBox.SelectedItem != null)
                {
                    var inboxFolder = (Outlook.MAPIFolder)this.FolderComboBox.SelectedItem;
                    ThisAddIn.thisApplication.ActiveExplorer().CurrentFolder = inboxFolder;
                    ThisAddIn.thisApplication.ActiveExplorer().CurrentFolder.Display();

                    var recivedDataList = this.GetReceived(inboxFolder);
                    this.InfoListBox.Items.Clear();
                    foreach (var recived in recivedDataList)
                    {
                        this.InfoListBox.Items.Add(recived);
                    }

                    this.InfoToolStripStatusLabel.Text = string.Format("Обнаружено: {0}", recivedDataList.Count);
                    MessageBox.Show(@"Recived скопированы в " + SavePath, @"Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                this.InfoListBox.Items.Add(MethodBase.GetCurrentMethod().Name + " Error message : " + ex.Message);
                this.InfoListBox.Items.Add(MethodBase.GetCurrentMethod().Name + " StackTrace " + ex.StackTrace);
            }
            finally
            {
                this.Cursor = currentCursor;
                Application.DoEvents();
            }
        }

        /// <summary>
        /// The save log button_ click.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private void SaveLogButtonClick(object sender, EventArgs e)
        {
            var savefile = new SaveFileDialog
                               {
                                   FileName = "received.txt",
                                   Filter = @"Text files (*.txt)|*.txt|All files (*.*)|*.*"
                               };

            if (savefile.ShowDialog() == DialogResult.OK)
            {
                using (var sw = new StreamWriter(savefile.FileName))
                {
                    foreach (var item in this.InfoListBox.Items)
                    {
                        sw.WriteLine(item);                        
                    }
                }
            }
        }

        /// <summary>
        /// The extract attachment.
        /// </summary>
        /// <param name="inboxFolder">
        /// The inbox folder.
        /// </param>
        /// <returns>
        /// The <see cref="List"/>.
        /// </returns>
        private List<string> ExtractAttachment(Outlook.MAPIFolder inboxFolder)
        {
            string fileName = string.Empty;
            if (!Directory.Exists(SavePath))
            {
                Directory.CreateDirectory(SavePath);
            }

            var mailBoxContent = new List<string>();
            if (inboxFolder == null)
            {
                return mailBoxContent;
            }

            var folderItems = inboxFolder.Items;
            folderItems.Sort("[CreationTime]", true);
            
            this.progressBar.Maximum = folderItems.Count;
            this.progressBar.Value = 0;

            foreach (object collectionItem in folderItems)
                {
                    try
                    {
                        var mail = collectionItem as Outlook.MailItem;
                        if (mail == null)
                        {
                            continue;
                        }

                        if (mail.Attachments.Count > 0)
                        {
                            for (var i = 1; i <= mail.Attachments.Count; i++)
                            {
                                var sender = this.RemoveIllegal(mail.Sender.Name);

                                fileName = Path.Combine(SavePath, sender);
                                if (!Directory.Exists(fileName))
                                {
                                    Directory.CreateDirectory(fileName);
                                }

                                var createDate = string.Format("{0:0000}.{1:00}.{2:00}", mail.CreationTime.Year, mail.CreationTime.Month, mail.CreationTime.Day);

                                if (mail.Subject != null)
                                {
                                    var subject = this.RemoveIllegal(mail.Subject);
                                    fileName = Path.Combine(SavePath, sender, createDate + " " + subject);
                                }
                                else
                                {
                                    fileName = Path.Combine(SavePath, sender, createDate);
                                }

                                if (!Directory.Exists(fileName))
                                {
                                    Directory.CreateDirectory(fileName);
                                }

                                fileName = Path.Combine(fileName, mail.Attachments[i].FileName);
                                mail.Attachments[i].SaveAsFile(fileName);
                                var msg = string.Format("{0} Subject: {1}    Attachment: {2}", mail.CreationTime.ToShortDateString(), mail.Subject, mail.Attachments[i].FileName);
                                mailBoxContent.Add(msg);
                            }
                        }

                        this.progressBar.Value++;
                        this.MainStatusStrip.Refresh();
                    }
                    catch (Exception ex)
                    {
                        this.InfoListBox.Items.Add(ex.Message);
                        this.InfoListBox.Items.Add(fileName);
                        if (this.errorList == null)
                        {
                            this.errorList = new List<string>();
                        }

                        this.errorList.Add(MethodBase.GetCurrentMethod().Name + " Error message : " + ex.Message);
                    }
            }

            this.progressBar.Value = 0;
            this.MainStatusStrip.Refresh();
            
            Marshal.ReleaseComObject(folderItems);
            return mailBoxContent;
        }

        /// <summary>
        /// The clear attachment.
        /// </summary>
        /// <param name="inboxFolder">
        /// The inbox folder.
        /// </param>
        /// <returns>
        /// The <see cref="List"/>.
        /// </returns>
        private List<string> ClearAttachment(Outlook.MAPIFolder inboxFolder)
        {
            var mailBoxContent = new List<string>();
            if (inboxFolder == null)
            {
                return mailBoxContent;
            }

            Outlook.Items folderItems = null;

            try
            {
                folderItems = inboxFolder.Items;
                folderItems.Sort("[CreationTime]", true);

                foreach (object collectionItem in folderItems)
                {
                    var newEmail = collectionItem as Outlook.MailItem;
                    if (newEmail != null)
                    {
                        if (newEmail.Attachments.Count > 0)
                        {
                            for (var i = 1; i <= newEmail.Attachments.Count; i++)
                            {
                                //if (newEmail.Attachments[i].FileName == "ATT00001.png")
                                {
                                    newEmail.Attachments[i].Delete();
                                }
                            }

                            newEmail.Save();
                        }
                    }
                }

                return mailBoxContent;
            }
            catch (Exception ex)
            {
                this.InfoListBox.Items.Add(ex.Message);
                if (this.errorList == null)
                {
                    this.errorList = new List<string>();
                }

                this.errorList.Add(MethodBase.GetCurrentMethod().Name + " Error message : " + ex.Message);
                this.errorList.Add(MethodBase.GetCurrentMethod().Name + " StackTrace " + ex.StackTrace);
                return null;
            }
            finally
            {
                if (folderItems != null)
                {
                    Marshal.ReleaseComObject(folderItems);
                }
            }
        }

        /// <summary>
        /// The get recived.
        /// </summary>
        /// <param name="inboxFolder">
        /// The inbox folder.
        /// </param>
        /// <returns>
        /// The <see cref="List"/>.
        /// </returns>
        private List<string> GetReceived(Outlook.MAPIFolder inboxFolder)
        {
            string fileName = string.Empty;
            if (!Directory.Exists(SavePath))
            {
                Directory.CreateDirectory(SavePath);
            }

            var receivedList = new List<string>();
            //receivedList.Add("ReceivedByName\tReceivedByEntryID\tReceivedOnBehalfOfName\tReceivedOnBehalfOfEntryID");
            if (inboxFolder == null)
            {
                return receivedList;
            }

            var folderItems = inboxFolder.Items;
            folderItems.Sort("[CreationTime]", true);

            this.progressBar.Maximum = folderItems.Count;
            this.progressBar.Value = 0;

            foreach (object collectionItem in folderItems)
            {
                try
                {
                    var mail = collectionItem as Outlook.MailItem;
                    if (mail == null)
                    {
                        continue;
                    }

                    var receivedByName = mail.ReceivedByName;
                    //var receivedByEntryID = mail.ReceivedByEntryID.Trim('\0');
                    //var receivedOnBehalfOfName = mail.ReceivedOnBehalfOfName;
                    //var receivedOnBehalfOfEntryID = mail.ReceivedOnBehalfOfEntryID.Trim('\0');
                    //var s = string.Format("{0}\t{1}\t{2}\t{3}", receivedByName, receivedByEntryID, receivedOnBehalfOfName, receivedOnBehalfOfEntryID);
                    var s = string.Format("{0}", receivedByName);
                    receivedList.Add(s);

                    this.progressBar.Value++;
                    this.MainStatusStrip.Refresh();
                }
                catch (Exception ex)
                {
                    this.InfoListBox.Items.Add(ex.Message);
                    this.InfoListBox.Items.Add(fileName);
                    if (this.errorList == null)
                    {
                        this.errorList = new List<string>();
                    }

                    this.errorList.Add(MethodBase.GetCurrentMethod().Name + " Error message : " + ex.Message);
                }
            }

            this.progressBar.Value = 0;
            this.MainStatusStrip.Refresh();

            Marshal.ReleaseComObject(folderItems);
            return receivedList;
        }

        /// <summary>
        /// The remove illegal.
        /// </summary>
        /// <param name="s">
        /// The s.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        private string RemoveIllegal(string s)
        {
            var invalid = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            foreach (var c in invalid)
            {
                s = s.Replace(c.ToString(), string.Empty);
            }

            return s.Trim();
        }

        /// <summary>
        /// The folder combo box_ selected index changed.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private void FolderComboBoxSelectedIndexChanged(object sender, EventArgs e)
        {
            this.InfoToolStripStatusLabel.Text = this.FolderComboBox.Text;
        }
    }
}
