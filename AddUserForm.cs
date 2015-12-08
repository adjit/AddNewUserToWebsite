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

namespace SendWebUsername
{
    public partial class AddUserForm : Form
    {
        public AddUserForm()
        {
            InitializeComponent();
        }

        private void sendButton_Click(object sender, EventArgs e)
        {
            Outlook.Application oApp = Globals.ThisAddIn.Application;
            Outlook.ContactItem oContact = (Outlook.ContactItem)oApp.CreateItem(Outlook.OlItemType.olContactItem);
            Outlook.MailItem oMail = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
            Outlook.NameSpace oNamespace = oApp.GetNamespace("MAPI");
            Outlook.MAPIFolder myContacts = oNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            Outlook.Items oItem = myContacts.Items;
            String filter = "[FirstName] = '" + firstName.Text + "' And [LastName] = '" + lastName.Text + "'";
            oItem = oItem.Find(filter);

            if(oItem != null)
            {
                DialogResult contactResult = System.Windows.Forms.MessageBox.Show("Contact already exists. Would you like to overwrite?", "Contact Exists", MessageBoxButtons.YesNo);
            }

            try
            {
                oContact.FirstName = firstName.Text;
                oContact.LastName = lastName.Text;
                oContact.Email1Address = email.Text;
                oContact.CompanyName = company.Text;
                oContact.Move(myContacts.Folders["Constant Contact"]);
                oContact.Save();
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("User Not Added to Contacts");
            }
        }
    }
}
