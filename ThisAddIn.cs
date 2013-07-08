using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using System.Net;
using System.IO;
using System.Xml;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;

namespace pptVivo2007Addin
{
    public partial class ThisAddIn
    {
        
        private string expositionId = "0";

        public static int userId = 0;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            //http://msdn.microsoft.com/en-us/library/vstudio/cc668192(v=vs.100).aspx
            //Slide change event handler            
            /*
            ((PowerPoint.EApplication_Event)this.Application).SlideSelectionChanged +=
                new Microsoft.Office.Interop.PowerPoint.
                    EApplication_SlideSelectionChangedEventHandler(
                    ThisAddIn_SlideSelectionChanged);
             */

            //Slide change event handler (funcioan igual que al anterior?)
            this.Application.SlideSelectionChanged +=
                new PowerPoint.EApplication_SlideSelectionChangedEventHandler(ThisAddIn_SlideSelectionChanged);

            //Slide change in slide show event handler
            this.Application.SlideShowNextSlide +=
                new PowerPoint.EApplication_SlideShowNextSlideEventHandler(ThisAddIn_SlideShowNextSlide);

            //After presentation open event handler
            this.Application.AfterPresentationOpen +=
                new PowerPoint.EApplication_AfterPresentationOpenEventHandler(ThisAddIn_AfterPresentationOpen);

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void LoadExpositionIdAfterLogin()
        {
            loadExpositionId(this.Application.ActivePresentation.Name);

            //Resets the DB exposition index
            if (!this.expositionId.Equals(0))
            {
                String slideId = "1";
                this.updateSlide(slideId);
            }

        }

        //Loads exposition id for current presentation
        void ThisAddIn_AfterPresentationOpen(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            if (!userId.Equals(0))
            {
                this.loadExpositionId(Pres.Name);

                //Resets the DB exposition index
                if (!this.expositionId.Equals(0))
                {
                    String slideId = "1";
                    this.updateSlide(slideId);
                }
            }
        }

        //Slide moving in slide show
        void ThisAddIn_SlideShowNextSlide(SlideShowWindow Wn)
        {
            if (!this.expositionId.Equals(0))
            {
                //add a text box to each new slide
                PowerPoint.Shape textBox = Wn.View.Slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
                textBox.TextFrame.TextRange.InsertAfter("Join us at: XXX/XXX");

                String slideId = Wn.View.Slide.SlideNumber.ToString();
                this.updateSlide(slideId);
            }
        }

        //Slide moving in edition view
        void ThisAddIn_SlideSelectionChanged(PowerPoint.SlideRange SldRange)
        {
            if (!this.expositionId.Equals(0))
            {
                String slideId = SldRange.SlideNumber.ToString();
                this.updateSlide(slideId);
            }
        }

        //Performs an http request to update selected slide
        void updateSlide(string slideId)
        {
            //Url to update slide
            string sURL;
            sURL = "http://www.pptvivo.com/services/updateSlide.php?action=updateSlide&slideId=" + slideId + "&expositionId=" + this.expositionId;

            //make call
            string webResponse;
            webResponse = this.makeHTTPRequest(sURL);

            //Xml Parsing
            //StringBuilder output = new StringBuilder();
            string result = "";

            using (XmlReader reader = XmlReader.Create(new StringReader(webResponse)))
            {
                reader.ReadToFollowing("response");
                reader.MoveToFirstAttribute();
                result = reader.Value;
                //output.AppendLine("Result: " + result);
            }

            if (result.Equals("false"))
            {
                MessageBox.Show("Error updating slide", "pptVivo!");
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

        //Loads the exposition id for current presentation
        void loadExpositionId(String presentationName)
        {
            //Url to get exposition id
            string sURL;
            sURL = "http://www.pptvivo.com/services/updateSlide.php?action=getExpositionId&authorId=" + userId + "&presentationName=" + presentationName;

            //make call
            string webResponse;
            webResponse = this.makeHTTPRequest(sURL);

            using (XmlReader reader = XmlReader.Create(new StringReader(webResponse)))
            {
                reader.ReadToFollowing("expositionId");
                reader.MoveToFirstAttribute();
                string result = reader.Value;
                this.expositionId = result;
            }

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
