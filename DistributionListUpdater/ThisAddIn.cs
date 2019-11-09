using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace DistributionListUpdater
{
    public partial class ThisAddIn
    {
        public List<DistListItem> DistLists = new List<DistListItem>();

        public Outlook.MAPIFolder ContactsFolder;
        public Outlook.MAPIFolder ListsFolder;

        private void ClearDistributionLists() {
            DistLists.Clear();
            foreach (ContactItem distList in ListsFolder.Items)
            {
                distList.Delete();
            }
        }

        /**
         * This is where the fun happens :-)
         */
        public void RebuildDistributionLists()
        {
            DistListItem myList = null;

            ClearDistributionLists();
            DlgWait StatusDlg = new DlgWait();
            StatusDlg.ShowDialog();

        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // ensure folder paths are set
            string ContactsFolderPath = Properties.Settings.Default.ContactsFolderPath;
            string ListsFolderPath = Properties.Settings.Default.ListsFolderPath;

            if (ContactsFolderPath == "")
            {
                ConfigureContactsFolder();
            } else
            {
                ContactsFolder = GetFolderByPath(ContactsFolderPath);
            }

            ListsFolder = GetFolderByPath(ListsFolderPath);
            while (null == ListsFolder) ConfigureListsFolder();


            ListsFolder.ShowAsOutlookAB = true;
        }


        public void ConfigureContactsFolder()
        {
            MessageBox.Show("Bitte wählen Sie im folgenden Dialog einen Ordner, dessen Kontakte als Grundlage für Verteilerlisten dienen sollen.", "Kein Kontaktordner angegeben");
            ContactsFolder = this.Application.GetNamespace("MAPI").PickFolder();
            Properties.Settings.Default.ContactsFolderPath = ContactsFolder.FolderPath;
            Properties.Settings.Default.Save();
        }

        public void ConfigureListsFolder()
        {
            MessageBox.Show("Bitte wählen Sie im folgenden Dialog einen Ordner, in dem die Verteilerlisten gespeichert werden sollen.", "Kein Listenordner angegeben");
            ListsFolder = this.Application.GetNamespace("MAPI").PickFolder();
            Properties.Settings.Default.ListsFolderPath = ListsFolder.FolderPath;
            Properties.Settings.Default.Save();
        }


        private Folder GetFolderByPath(string path)
        {
            foreach (Store store in Application.Session.Stores)
            {
                Folder folder = GetSubFolderByPath((Folder)store.GetRootFolder(), path);
                if (null != folder) return folder;
            }
            return null;
        }

        private Folder GetSubFolderByPath(Folder folder, string path)
        {

            if (path == folder.FolderPath) return folder;
            foreach (Folder subFolder in folder.Folders)
            {
                Folder myFolder = GetSubFolderByPath(subFolder, path);
                if (null != myFolder) return myFolder;
            }
            return null;

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }


        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
