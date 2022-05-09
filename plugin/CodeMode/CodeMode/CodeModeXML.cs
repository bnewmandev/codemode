using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new CodeModeXML();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace CodeMode
{

    public class FontItem : RibbonDropDownItem
    {
        public string Label { get; set; }
        public string Id { get; }
        public Image Image { get; set; }
        public string OfficeImageId { get; set; }
        public OfficeRibbon Ribbon { get; }
        public RibbonComponent Parent { get; }
        public string ScreenTip { get; set; }
        public string SuperTip { get; set; }
        public object Tag { get; set; }
    }


    [ComVisible(true)]
    public class CodeModeXML : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        ColorDialog colorPicker;
        static WdColor codeColor;
        static WdColor bgColor;

        public CodeModeXML()
        {

        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("CodeMode.CodeModeXML.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226


        public void Ribbon_Load(Office.IRibbonControl control)
        {
            foreach (FontFamily font in FontFamily.Families)
            {
                Globals.ThisAddIn.listItems.Append(font.Name);
            }
        }

        public void Load(Office.IRibbonControl control)
        {
            foreach (FontFamily font in FontFamily.Families)
            {
                Globals.ThisAddIn.listItems.Append(font.Name);
            }
        }

        public int getCount(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.listItems.Count;
        }


        public string getLabel(Office.IRibbonControl control, int index)
        {
            return Globals.ThisAddIn.listItems[index];
        }



        public void toggleCodeMode_Click(Office.IRibbonControl control, bool isPressed)
        {
            Debug.WriteLine(Globals.ThisAddIn.listItems[0]);
            Debug.WriteLine(control);
            if (isPressed)
            {
                Selection initialLocation = Globals.ThisAddIn.Application.Selection;
                ThisAddIn.initialLocation = initialLocation.Range;

                string initialFont = initialLocation.Font.Name;
                ThisAddIn.initialFont = initialFont;

                initialLocation.Font.Name = "Courier New";


                Debug.WriteLine(initialLocation.Start);
            }
            else
            {
                Selection finalLocation = Globals.ThisAddIn.Application.Selection;
                Range fullCodeRange = Globals.ThisAddIn.Application.ActiveDocument.Range(ThisAddIn.initialLocation.Start, finalLocation.End);
                fullCodeRange.Select();
                fullCodeRange.NoProofing = 1;
                fullCodeRange.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic;
                fullCodeRange.Shading.BackgroundPatternColor = (WdColor)2829099;
                Selection selection = Globals.ThisAddIn.Application.Selection;
                selection.MoveRight();
                selection.InsertAfter(" ");
                selection.Font.Name = ThisAddIn.initialFont;
                selection.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic;
                selection.MoveRight();
            }
        }


        public void selectCodeColor_Click(Office.IRibbonControl control)
        {
            colorPicker = new ColorDialog();
            colorPicker.AllowFullOpen = true;
            colorPicker.ShowHelp = true;
            colorPicker.ShowDialog();

            string colorString = colorPicker.Color.R.ToString() + colorPicker.Color.G.ToString() + colorPicker.Color.B.ToString();
            codeColor = (WdColor)Int32.Parse(colorString);

        }

        public void selectBackgroundColor_Click(Office.IRibbonControl control)
        {
            colorPicker = new ColorDialog();
            colorPicker.AllowFullOpen = true;
            colorPicker.ShowHelp = true;
            colorPicker.ShowDialog();

            string colorString = colorPicker.Color.R.ToString() + colorPicker.Color.G.ToString() + colorPicker.Color.B.ToString();
            bgColor = (WdColor)Int32.Parse(colorString);
        }

        public void codeFont_TextChanged(Office.IRibbonControl control)
        {

        }



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
