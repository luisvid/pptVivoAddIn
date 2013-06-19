using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace pptVivo2007Addin
{
    public partial class pptVivoRibbon
    {
        private void pptVivoRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnLogin_Click(object sender, RibbonControlEventArgs e)
        {

            //Open login form
            if (pptVivo2007Addin.ThisAddIn.userId.Equals(0))
            {
                frmLogin objChild = frmLogin.GetChild();
                objChild.Show();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("User already logged", "pptVivo!");
            }
        }

        private void btnAccount_Click(object sender, RibbonControlEventArgs e)
        {
            //System.Windows.Forms.MessageBox.Show("Go to Website", "pptVivo!");
            //Process.Start(http://myserver/path/file.aspx?para1=data&para2=more+data);
            //Use HttpUtility.UrlEncode to encode strings as parameters.

            OpenLink("http://www.pptvivo.com");
        }

        public void OpenLink(string sUrl)
        {
            try
            {
                System.Diagnostics.Process.Start(sUrl);
            }
            catch (Exception exc1)
            {
                // System.ComponentModel.Win32Exception is a known exception that occurs when Firefox is default browser.            // It actually opens the browser but STILL throws this exception so we can just ignore it.  If not this exception,
                // then attempt to open the URL in IE instead.
                if (exc1.GetType().ToString() != "System.ComponentModel.Win32Exception")
                {
                    // sometimes throws exception so we have to just ignore
                    // this is a common .NET bug that no one online really has a great reason for so now we just need to try to open
                    // the URL using IE if we can.
                    try
                    {
                        System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo("IExplore.exe", sUrl);
                        System.Diagnostics.Process.Start(startInfo);
                        startInfo = null;
                    }
                    catch (Exception exc2)
                    {
                        // still nothing we can do so just show the error to the user here.
                    }
                }
            }
        }

        private void btnHelp_Click(object sender, RibbonControlEventArgs e)
        {
       
        }

        private void btnHelp_Click_1(object sender, RibbonControlEventArgs e)
        {
            OpenLink("http://pptvivo.com/en?action=downloadInfo");
        }
    }
}
