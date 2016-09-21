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

        private void CheckSettings()
        {
            // Check Settings and insert some default values if null (first run) then show the Settings dialog.
            if (String.IsNullOrEmpty(Properties.Settings.Default.Host)
                || String.IsNullOrEmpty(Properties.Settings.Default.Account)
                || String.IsNullOrEmpty(Properties.Settings.Default.Password)
                || String.IsNullOrEmpty(Properties.Settings.Default.Contract))
            {
                Properties.Settings.Default.Host = "http://localhost:8545";
                Properties.Settings.Default.Account = "0x1234";
                Properties.Settings.Default.Password = "";
                Properties.Settings.Default.Contract = "0xabcd";

                // Show settings
                SettingsDialog settingsDialog = new SettingsDialog();
                settingsDialog.Show();
            }
        }

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
            // Check if Settings are OK
            CheckSettings();
            
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
            // Check if Settings are OK
            CheckSettings();

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

        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            SettingsDialog settingsDialog = new SettingsDialog();
            settingsDialog.Show();
        }
    }
}
