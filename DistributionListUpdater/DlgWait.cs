using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DistributionListUpdater
{
    public partial class DlgWait : Form
    {
        public DlgWait()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        public void SetMax(int n)
        {
            this.progressBar1.Maximum = n;
        }

        public void SetValue(int n)
        {
            this.progressBar1.Value = n;
        }

        private void DlgWait_Load(object sender, EventArgs e) {
            this.SetMax(Globals.ThisAddIn.ContactsFolder.Items.OfType<Outlook.ContactItem>().Count());
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            backgroundWorker1.RunWorkerAsync();
        }

        private void DlgWait_Shown(object sender, EventArgs e)
        {

        }

        //call back method
        public void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            Outlook.DistListItem myList;
            foreach (Outlook.DistListItem distList in Globals.ThisAddIn.ListsFolder.Items)
            {
                distList.Delete();
            }

            int max = Globals.ThisAddIn.ContactsFolder.Items.OfType<Outlook.ContactItem>().Count();

            int ctr = 0;
            foreach (Outlook.ContactItem contact in Globals.ThisAddIn.ContactsFolder.Items.OfType<Outlook.ContactItem>())
            {
                ctr++;
                backgroundWorker1.ReportProgress(ctr);
                Application.DoEvents();

                // only process contacts with an email address
                if (contact.Email1Address != null)
                {
                    // create a Recipient for this contact
                    Outlook.Recipient myRecipient = Globals.ThisAddIn.Application.Session.CreateRecipient(contact.Email1Address.ToString());
                    myRecipient.Resolve();

                    // add to relevant DistListItems according to categories
                    if (contact.Categories != null)
                    {
                        var categories = contact.Categories.Split(new[] { "; " }, StringSplitOptions.None);
                        foreach (string category in categories)
                        {
                            if (!Globals.ThisAddIn.DistLists.Exists(x => x.DLName == "VL." + category))
                            {
                                // DistListItem needs to be created
                                myList = Globals.ThisAddIn.ListsFolder.Items.Add(Outlook.OlItemType.olDistributionListItem) as Outlook.DistListItem;
                                myList.DLName = "VL." + category;
                                myList.Body = "Dies ist eine automatisch erzeugte Liste. Sie sollte nicht von Hand geändert werden, da beim nächsten Update alle Änderungen überschrieben würden.";
                                myList.Save();
                                Globals.ThisAddIn.DistLists.Add(myList);
                            }
                            else
                            {
                                // DistListItem is already present
                                myList = Globals.ThisAddIn.DistLists.Find(x => x.DLName == "VL." + category);
                            }
                            myList.AddMember(myRecipient);
                            myList.Save();
                        }
                    }
                }
            }



        }
        public void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = progressBar1.Maximum;
            this.Hide();
        }


    }
}
