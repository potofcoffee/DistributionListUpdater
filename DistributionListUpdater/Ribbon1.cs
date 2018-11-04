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
    }
}
