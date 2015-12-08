using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace SendWebUsername
{
    public partial class AddUserRibbon
    {
        private void AddUserRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void addUserButton_Click(object sender, RibbonControlEventArgs e)
        {
            AddUserForm oForm = new AddUserForm();
            oForm.Show();
        }
    }
}
