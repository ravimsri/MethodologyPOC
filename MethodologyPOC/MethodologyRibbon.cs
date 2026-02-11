using System;
using Microsoft.Office.Tools.Ribbon;

namespace MethodologyPOC
{
    public partial class MethodologyRibbon
    {

        private void MethodologyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            ApplyVisibility();
        }

        public void ApplyVisibility()
        {
            var email = (ThisAddIn.CurrentEmail ?? "").Trim();
            btnHelloWorld.Visible = email.Equals(
                ConfigHelper.User1Email,
                StringComparison.OrdinalIgnoreCase);

        }

        private void btnHelloWorld_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Hello World");
        }
    }
}
