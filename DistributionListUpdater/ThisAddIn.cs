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
        private List<MAPIFolder> ContactFolders = new List<MAPIFolder>();
        private MAPIFolder DistListFolder;
        private List<DistListItem> DistLists = new List<DistListItem>();


        private void GetSubFolders(MAPIFolder folder) {
            foreach (MAPIFolder subFolder in folder.Folders)
            {
                if (subFolder.DefaultMessageClass == "IPM.Contact") this.ContactFolders.Add(subFolder);
                if (subFolder.Folders.Count > 0) GetSubFolders(subFolder);
            }
        }


        private void GetAllContactFolders()
        {
            foreach (Store store in Application.Session.Stores)
            {
                GetSubFolders(store.GetRootFolder());
            }
            // exempt the DistListFolder from this list
            ContactFolders.Remove(DistListFolder);
        }

        private void ClearDistributionLists() {
            DistLists.Clear();
            foreach (DistListItem distList in DistListFolder.Items.OfType<DistListItem>())
            {
                for (int i=1; i<=distList.MemberCount; i++)
                {
                    distList.RemoveMember(distList.GetMember(1));
                }
                DistLists.Add(distList);
            }
        }

        /**
         * This is where the fun happens :-)
         */
        public void RebuildDistributionLists()
        {
            DistListItem myList = null;

            EnsureDistListFolderIsPresent();
            ClearDistributionLists();

            // process all contact folders
            foreach (MAPIFolder currentFolder in ContactFolders)
            {
                // repeat for each contact in this folder
                foreach (ContactItem contact in currentFolder.Items.OfType<ContactItem>())
                {
                    // only process contacts with an email address
                    if (contact.Email1Address != null)
                    {
                        // create a Recipient for this contact
                        Recipient myRecipient = Application.Session.CreateRecipient(contact.Email1Address.ToString());
                        myRecipient.Resolve();

                        // add to relevant DistListItems according to categories
                        if (contact.Categories != null)
                        {
                            var categories = contact.Categories.Split(new[] { "; " }, StringSplitOptions.None);
                            foreach (string category in categories)
                            {
                                if (!DistLists.Exists(x => x.DLName == "VL." + category))
                                {
                                    // DistListItem needs to be created
                                    myList = DistListFolder.Items.Add(OlItemType.olDistributionListItem) as DistListItem;
                                    myList.DLName = "VL." + category;
                                    myList.Body = "Dies ist eine automatisch erzeugte Liste. Sie sollte nicht von Hand geändert werden, da beim nächsten Update alle Änderungen überschrieben würden.";
                                    myList.Save();
                                    DistLists.Add(myList);
                                }
                                else
                                {
                                    // DistListItem is already present
                                    myList = DistLists.Find(x => x.DLName == "VL." + category);
                                }
                                myList.AddMember(myRecipient);
                                myList.Save();
                            }
                        }
                    }
                }
            }
        }

        private void EnsureDistListFolderIsPresent()
        {
            var defaultContactsFolder = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            try
            {
                DistListFolder = defaultContactsFolder.Folders["Verteilerlisten"];
            }
            catch
            {
                defaultContactsFolder.Folders.Add("Verteilerlisten");
                DistListFolder = defaultContactsFolder.Folders["Verteilerlisten"];
            }
            this.DistListFolder.ShowAsOutlookAB = true;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // ensure the distribution list folder is present
            EnsureDistListFolderIsPresent();

            // make a list of all contacts folders
            GetAllContactFolders();
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
