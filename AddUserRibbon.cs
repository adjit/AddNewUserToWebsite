using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SendWebUsername
{
    public partial class AddUserRibbon
    {
        private void AddUserRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void addUserButton_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application oApp = Globals.ThisAddIn.Application;
            Outlook.ContactItem oContact = (Outlook.ContactItem)oApp.CreateItem(Outlook.OlItemType.olContactItem);


            try
            {

            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("User Not Added to Contacts");
            }
        }
    }
}
