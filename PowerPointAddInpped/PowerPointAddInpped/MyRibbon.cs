using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.Drawing;


// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MyRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PowerPointAddInpped
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {

        public static string s = "LiveSlide \r\n http://rd.gallowayridge.com/Frame.html";
        public void OnTextButton(Office.IRibbonControl control)
        {
            //System.Drawing.Point position = new System.Drawing.Point();
            //MonthCalendar Month = new MonthCalendar()
            SelectForm();
           // CreateMyForm();
        }

     
        public void SelectForm ()
        {
            //Definess the form 
            Form Form1 = new Form();
            //Defines the buttons and labels
             Button button1 = new Button();
            TextBox Index = new TextBox();
            Label SN = new Label();
            MonthCalendar Month1 = new MonthCalendar();
            Label CurrentSlide = new Label();
            DateTimePicker timePicker = new DateTimePicker();

            //Delfines the Text for each button or label
            Form1.Text = "Select Date";
            button1.Text = "Expire Slide";
            SN.Text = "Slide Number:";
            SN.Location = new System.Drawing.Point(10,10);
            Index.Location = new System.Drawing.Point(15, 25);

            
            button1.Location = new System.Drawing.Point(180, 220);
            Form1.AcceptButton = button1;
            button1.DialogResult = DialogResult.Yes;

            Form1.Controls.Add(button1);   
            Form1.Controls.Add(Index);
            Form1.Controls.Add(SN);
            Form1.Size = new Size(300, 300);
            Form1.FormBorderStyle = FormBorderStyle.FixedDialog;
            Form1.MaximizeBox = false;
           
            timePicker.Format = DateTimePickerFormat.Time;
            timePicker.ShowUpDown = true;
            timePicker.Location = new System.Drawing.Point(130, 25);
            timePicker.Width = 100;
            timePicker.CustomFormat = "HH:mm";
            Form1.Controls.Add(timePicker);


            Globals.ThisAddIn.Application.ActiveWindow.Panes[2].Activate();
            var currentSlideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex;
            string curr = Convert.ToString(currentSlideIndex);
           
            Index.Text = curr;


            
            Month1.Location = new System.Drawing.Point(15, 45);
            Form1.Controls.Add(Month1);

            var presentation = Globals.ThisAddIn.Application;
            
           
            CurrentSlide.Location = new System.Drawing.Point(15, 15);
            Form1.ShowDialog();



            if (Form1.DialogResult == DialogResult.Yes)
            {
                int number = presentation.ActivePresentation.Slides.Count;
                string number1 = Convert.ToString(number);
                string numberentry = Convert.ToString(currentSlideIndex);
             
               
                var Date = Month1.SelectionRange.Start.ToString("MM/dd/yyyy");
                curr = Index.Text;
                bool number2 = int.TryParse(curr, out int num);
                if (number2 == false)
                {
                    MessageBox.Show("Please enter a number. Not a letter or symbol");
                    SelectForm();
                }
                int slidenumber = Convert.ToInt32(curr);
                var T = timePicker.Value.ToString("HH:mm tt");
                float n = 5;
                float p = 5;
                if (slidenumber <= number)
                {
                    //  MessageBox.Show(curr);
                    //  MessageBox.Show(Date);
                    // MessageBox.Show(T);
                    //Globals.ThisAddIn.Application.ActiveWindow.Panes[2].Activate();
                    // presentation.ActivePresentation.Slides[slidenumber].Name = Date;
                    // string s = presentation.ActivePresentation.Slides[slidenumber].NotesPage.Shapes[2].TextFrame.TextRange.Text;
                    presentation.ActivePresentation.Slides[slidenumber].NotesPage.Shapes[2].TextFrame.TextRange.Text =  Date + " " + T;


                //presentation.ActivePresentation.Slides[slidenumber].Comments.Add(n,p, "KMS", "KMS", Date + " " + T);
                //  presentation.ActivePresentation.Slides[slidenumber].SlideShowTransition.Hidden = Office.MsoTriState.msoTrue;
            }
                else
                {
                    MessageBox.Show("That slide does not exist please enter a valid slide number");
                    SelectForm();
                }
            }
            
        }

        public void OnTextButton1(Office.IRibbonControl control)
        {
            //MessageBox.Show("in Toggle");
            Globals.ThisAddIn.Application.ActiveWindow.Panes[2].Activate();
            var currentSlideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex;
            var presentation = Globals.ThisAddIn.Application;
            string curr = Convert.ToString(currentSlideIndex);
            int slidenumber = Convert.ToInt32(curr);
            presentation.ActivePresentation.Slides[slidenumber].SlideShowTransition.Hidden = Office.MsoTriState.msoTriStateToggle;
            
        }
        public void OnTextButton0(Office.IRibbonControl control)
        {

            //MessageBox.Show("in void");
            Globals.ThisAddIn.Application.ActiveWindow.Panes[2].Activate();
            var currentSlideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex;
            var presentation = Globals.ThisAddIn.Application;
            string curr = Convert.ToString(currentSlideIndex);
            int slidenumber = Convert.ToInt32(curr);
            var Hidden = presentation.ActivePresentation.Slides[slidenumber].SlideShowTransition.Hidden;
            if (Hidden == Office.MsoTriState.msoFalse)
            {
              //  MessageBox.Show("not hidden Yet");
                presentation.ActivePresentation.Slides[slidenumber].NotesPage.Shapes[2].TextFrame.DeleteText();
                presentation.ActivePresentation.Slides[slidenumber].NotesPage.Shapes[2].TextFrame.TextRange.Text = s;
            }
            else
            {
                presentation.ActivePresentation.Slides[slidenumber].SlideShowTransition.Hidden = Office.MsoTriState.msoFalse;
                presentation.ActivePresentation.Slides[slidenumber].NotesPage.Shapes[2].TextFrame.TextRange.Text = s;
            }

        }



        private Office.IRibbonUI ribbon;
    
        public void button1_Click(object sender, System.EventArgs e)
        {
            
           
         
        }
        public MyRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PowerPointAddInpped.MyRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
