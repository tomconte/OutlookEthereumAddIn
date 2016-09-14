using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Newtonsoft.Json;
using System.Security.Cryptography;
using System.IO;
using Nethereum.Hex.HexTypes;

namespace OutlookEthereumAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var inspectors = this.Application.Inspectors;
            //inspectors.NewInspector += Inspectors_NewInspector;
            this.Application.ItemSend += Application_ItemSend;
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            var item = Item as Outlook.MailItem;

            // Get the item properties (that were set by the ribbon check boxes)
            var isBlockchain = item.UserProperties == null || item.UserProperties["BlockchainStamp"] == null ? false : item.UserProperties["BlockchainStamp"].Value;
            var isNotify = item.UserProperties == null || item.UserProperties["BlockchainNotify"] == null ? false : item.UserProperties["BlockchainNotify"].Value;

            // If yes, calculate the hash and send it to the Smart Contract
            if (isBlockchain)
            {
                // Temporarily save the full item to disk and read back the contents
                var tempFile = Path.GetTempFileName();
                item.SaveAs(tempFile, Outlook.OlSaveAsType.olMSG);
                var itemMsgBytes = File.ReadAllBytes(tempFile);

                // Serialize item to JSON
                //var itemJson = JsonConvert.SerializeObject(itemMsg);

                // Calculate SHA256 hash
                //byte[] bytes = Encoding.UTF8.GetBytes(itemMsg);
                SHA256Managed sha256 = new SHA256Managed();
                byte[] hashbytes = sha256.ComputeHash(itemMsgBytes);
                StringBuilder hashstring = new StringBuilder("0x");
                foreach (Byte b in hashbytes)
                    hashstring.Append(b.ToString("x2"));
                var hash = hashstring.ToString();

                // Get message properties
                var subject = item.Subject;
                var sender = Application.Session.Accounts[1].SmtpAddress; // Use default account
                var itemRecipients = item.Recipients;
                StringBuilder recipientsstring = new StringBuilder();
                foreach (Outlook.Recipient r in itemRecipients)
                {
                    recipientsstring.Append(r.Address);
                    if (r.Index < itemRecipients.Count-1)
                        recipientsstring.Append(";");
                }
                var recipients = recipientsstring.ToString();

                // Blockchain hash
                var web3 = new Nethereum.Web3.Web3("http://cryptohash.tech:8545/");
                //var account = "0x87b3f6def4d451c41be733b8924da66dea0caed4"; // WARN: Hard-coded dev account
                var account = "0xca927db4582643f6ff2cf9606815eb5896ea361c";

                // Unlock account
                // Need to expose the personal API:
                // start geth --datadir Ethereum-Private --networkid 42 --nodiscover --rpc --rpcapi eth,web3,personal --rpccorsdomain "*" console
                web3.Personal.UnlockAccount.SendRequestAsync(account, "Pass123!", new HexBigInteger(120)).Wait();

                // Send transaction
                var abi = @"[{ ""constant"":false,""inputs"":[{""name"":""hash"",""type"":""uint256""},{""name"":""path"",""type"":""string""},{""name"":""computer"",""type"":""string""}],""name"":""fossilizeDocument"",""outputs"":[],""type"":""function""},{""constant"":true,""inputs"":[{""name"":"""",""type"":""uint256""}],""name"":""emails"",""outputs"":[{""name"":""sender"",""type"":""address""},{""name"":""subject"",""type"":""string""},{""name"":""emailFrom"",""type"":""string""},{""name"":""emailTo"",""type"":""string""}],""type"":""function""},{""constant"":false,""inputs"":[{""name"":""hash"",""type"":""uint256""},{""name"":""subject"",""type"":""string""},{""name"":""emailFrom"",""type"":""string""},{""name"":""emailTo"",""type"":""string""}],""name"":""fossilizeEmail"",""outputs"":[],""type"":""function""},{""constant"":true,""inputs"":[{""name"":"""",""type"":""uint256""}],""name"":""documents"",""outputs"":[{""name"":""sender"",""type"":""address""},{""name"":""path"",""type"":""string""},{""name"":""computer"",""type"":""string""}],""type"":""function""},{""anonymous"":false,""inputs"":[{""indexed"":false,""name"":""timestamp"",""type"":""uint256""},{""indexed"":true,""name"":""sender"",""type"":""address""},{""indexed"":false,""name"":""path"",""type"":""string""},{""indexed"":false,""name"":""computer"",""type"":""string""}],""name"":""DocumentFossilized"",""type"":""event""},{""anonymous"":false,""inputs"":[{""indexed"":false,""name"":""timestamp"",""type"":""uint256""},{""indexed"":true,""name"":""sender"",""type"":""address""},{""indexed"":false,""name"":""subject"",""type"":""string""},{""indexed"":false,""name"":""emailFrom"",""type"":""string""},{""indexed"":false,""name"":""emailTo"",""type"":""string""}],""name"":""EmailFossilized"",""type"":""event""}]";
                var address = "0x8bb0bfb8a3b7776a0646d96184f3355b1ee435f6";
                var contract = web3.Eth.GetContract(abi, address);
                var fossilizeFunc = contract.GetFunction("fossilizeEmail");
                fossilizeFunc.SendTransactionAsync(account, new HexBigInteger(1000000), new HexBigInteger(0), hash, subject, sender, recipients).Wait();

                // Send a confirmation e-mail

                Outlook.MailItem newItem = (Outlook.MailItem)this.Application.CreateItem(Outlook.OlItemType.olMailItem);
                newItem.Subject = "Message stamp confirmation " + hash;
                newItem.Body = "This is a confirmation that the attached message has been stamped in the blockchain with hash " + hash + ".";
                newItem.Attachments.Add(tempFile, Outlook.OlAttachmentType.olEmbeddeditem);

                // Do we need to notify the recipients?
                if (isNotify)
                {
                    StringBuilder notifyTo = new StringBuilder(sender);
                    notifyTo.Append(";");
                    // Also include recipients in the confirmation message
                    foreach (Outlook.Recipient r in itemRecipients)
                    {
                        notifyTo.Append(r.Address);
                        if (r.Index < itemRecipients.Count - 1)
                            notifyTo.Append(";");
                    }
                } else {
                    newItem.To = sender;
                }

                // Send the message
                newItem.Send();
            }
        }

        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            throw new NotImplementedException();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
