using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace KioskClient1
{
    public partial class ThisAddIn
    {
        void SlideShowBegin()
        {
            
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //I dont think slideshowbegin is nessecary 
            Application.SlideShowBegin += Application_SlideShowBegin;
            //When the nest slide is shown go to Application_SlideShowNextSlide
            Application.SlideShowNextSlide += Application_SlideShowNextSlide;
        }

        private void Application_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {

            // I dont think  this is nessecary
            int totalslides = Application.ActivePresentation.Slides.Count;
            int m = Application.SlideShowWindows[1].View.Slide.SlideIndex;
            string totalslide = Convert.ToString(totalslides);


            var presentation = Globals.ThisAddIn.Application;
            //get the current slide that's on the screen 
            int currentslide = Application.SlideShowWindows[1].View.Slide.SlideIndex;
            //get the notes for that current slide 
            string Notes = presentation.ActivePresentation.Slides[currentslide].NotesPage.Shapes[2].TextFrame.TextRange.Text;
            //Get the last line in the slide which will be the date time the user wants the slide to disappear
            string tail = Notes.Substring(Notes.LastIndexOf('\r') + 1);
            // sets the Provide for parsing the Date / Time
            System.Globalization.CultureInfo provider = new System.Globalization.CultureInfo("en-US");
            // Sets the input format for the expected Date/Time. 
            string inputFormat = "MM/dd/yyyy HH:mm tt";
            // If the Date / Time read in the notes is nothing
            if (tail == "")
            { 
                //do nothing
            }
            else
            {
                //else  convert the string into a Date/Time value
                DateTime resultDate = DateTime.ParseExact(tail, inputFormat, provider.DateTimeFormat);
                //if the read Date/Time is before the Date/Time right now
                if (resultDate <= DateTime.Now)
                {
                    //Delete the slide 
                    presentation.ActivePresentation.Slides[currentslide].Delete();
                       
                }
            }
            //not nessecary
            string current = Convert.ToString(currentslide);
           
        }
           

            private void Application_SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {

           //Not nessecary
            int totalslides = Application.ActivePresentation.Slides.Count;
            int currentslide = Application.SlideShowWindows[1].View.Slide.SlideIndex;
            string totalslide = Convert.ToString(totalslides);
           

        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
