using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;

namespace OutlookEthereumAddIn
{
    public partial class RibbonStamp
    {
        private void RibbonStamp_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
            // Get check box state
            var checkbox = sender as RibbonCheckBox;
            var state = checkbox.Checked;
            // Get current mail item
            var context = e.Control.Context as Inspector;
            var item = context.CurrentItem as MailItem;
            // Create User Property, and add as custom field
            item.UserProperties.Add("BlockchainStamp", OlUserPropertyType.olYesNo, true);
            // Set value to true
            item.UserProperties["BlockchainStamp"].Value = state;
        }

        private void checkBox2_Click(object sender, RibbonControlEventArgs e)
        {
            // Get check box state
            var checkbox = sender as RibbonCheckBox;
            var state = checkbox.Checked;
            // Get current mail item
            var context = e.Control.Context as Inspector;
            var item = context.CurrentItem as MailItem;
            // Create User Property, and add as custom field
            item.UserProperties.Add("BlockchainNotify", OlUserPropertyType.olYesNo, true);
            // Set value to true
            item.UserProperties["BlockchainNotify"].Value = state;
        }
    }
}
