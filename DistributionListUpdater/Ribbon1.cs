using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using DistributionListUpdater;

namespace DistributionListUpdater
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnUpdateAllLists_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RebuildDistributionLists();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RebuildDistributionLists();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ConfigureContactsFolder();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ConfigureContactsFolder();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ConfigureListsFolder();
        }

        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ConfigureContactsFolder();
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ConfigureListsFolder();
        }
    }
}
