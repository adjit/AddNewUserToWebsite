using System;
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

            String fName = firstName.Text;
            String lName = lastName.Text;
            String eMail = email.Text;
            String companyName = company.Text;
            String sUsername = username.Text;
            String sPassword = password.Text;
            String rep = repEmail.Text;

            String filter = "[FirstName] = '" + firstName.Text + "' And [LastName] = '" + lastName.Text + "'";

            String signature, mailBody;

            oMail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            oMail.Display();
            signature = oMail.HTMLBody;

            mailBody = "<HTML><BODY><P>Hi " + firstName.Text + ",</P>"
                + "<P>Thank you for your interest in the <STRONG>Metropolitan Sales</STRONG> website.</P>"
                + "<P>Some of the features of the website include:</P>"
                + "<UL><LI>Placing Orders</LI><LI>Order status & tracking</LI><LI>Detailed product information</LI>"
                + "<LI>Specification sheets in PDF for all products</LI></UL>"
                + "<P>These features can be accessed at:</P>"
                + "<P><a href= 'https://www.metsales.com'>www.metsales.com</a>, then click on Catalog</p>"
                + "<p><strong>Username : </strong>" + sUsername + "<br>"
                + "<strong>Password  : </strong>" + sPassword + "</p>"
                + "<p>Feel free to contact me should you have any questions.</p>"
                + "<p>Thank you," + signature + "</p></body></html>";

            /*
            
            ---- Filter not functioning as desired - does not find duplicate entries ----

            oItem = oItem.Find(filter);

            if(oItem != null)
            {
                DialogResult contactResult = 
                    System.Windows.Forms.MessageBox.Show("Contact already exists. Would you like to overwrite?", "Contact Exists", MessageBoxButtons.YesNo);
                if(contactResult == DialogResult.Yes) oContact = oItem as Outlook.ContactItem;

            }*/

            try
            {
                oContact.FirstName = fName;
                oContact.LastName = lName;
                oContact.Email1Address = eMail;
                oContact.CompanyName = companyName;
                oContact.Move(myContacts.Folders["Constant Contact"]);
                oContact.Save();
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("User Not Added to Contacts");
            }

            try
            {
                oMail.Recipients.Add(eMail);
                oMail.CC = rep;
                oMail.Subject = fName + " " + lName + " - Metropolitan Sales Username and Password";
                oMail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                oMail.HTMLBody = mailBody;
                oMail.Send();
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Did not send email to user");
            }

            this.Dispose();
        }
    }
}
