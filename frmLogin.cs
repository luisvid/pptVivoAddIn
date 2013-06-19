using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Xml;

namespace pptVivo2007Addin
{
    public partial class frmLogin : Form
    {

        private Int32 userId;

        private static frmLogin _childInstance = null;

        public static frmLogin GetChild()
        {
            //here I check for any opened instances
            //if they are null I create one
            //otherwise I just set the focus
            //on the existing form
            if (_childInstance == null)
            {
                _childInstance = new frmLogin();
            }
            else
            {
                _childInstance.Focus();
            }

            return _childInstance;
        }
        

        public frmLogin()
        {
            InitializeComponent();
        }

  
        private void cmdLogin_Click(object sender, EventArgs e)
        {

            String userLogin = this.txtUsername.Text;
            String userPassword = this.txtPassword.Text;

            //Url to login
            string sURL;
            sURL = "http://www.pptvivo.com/services/login.php?userLogin=" + userLogin + "&userPassword=" + userPassword;

            //make call
            string webResponse;
            webResponse = this.makeHTTPRequest(sURL);

            //Xml Parsing
            string result = "";
            string resultMessage = "";

            using (XmlReader reader = XmlReader.Create(new StringReader(webResponse)))
            {
                reader.ReadToFollowing("response");
                reader.MoveToFirstAttribute();
                result = reader.Value;
                reader.MoveToAttribute("message");
                resultMessage = reader.Value;
            }

            this.userId = Convert.ToInt32(result);

            if (this.userId.Equals(0))
            {
                MessageBox.Show(resultMessage, "pptVivo! Login");
            }
            else
            {
                pptVivo2007Addin.ThisAddIn.userId = this.userId;
                MessageBox.Show("Login successful. Welcome " + userLogin, "pptVivo! Login");
                this.Close();
                _childInstance = null;
            }
        }

        //Performs the http request
        String makeHTTPRequest(String sURL)
        {
            //Web request
            WebRequest wrGETURL;
            wrGETURL = WebRequest.Create(sURL);

            Stream objStream;
            objStream = wrGETURL.GetResponse().GetResponseStream();

            //Read http response
            StreamReader objReader = new StreamReader(objStream);
            String webResponse = objReader.ReadToEnd();

            return webResponse;
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {

        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            _childInstance = null;
        }

    }
}
