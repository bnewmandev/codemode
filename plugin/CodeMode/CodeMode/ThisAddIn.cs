using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Drawing;
using System.Drawing.Text;
using Microsoft.Office.Interop.Word;

namespace CodeMode
{

    public partial class ThisAddIn
    {
        public List<string> listItems = new List<string>();
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            InstalledFontCollection installedFontCollection = new InstalledFontCollection();

            foreach (FontFamily font in installedFontCollection.Families)
            {
                listItems.Add(font.Name);
            }
            return new CodeModeXML();
        }

        public static Word.Range initialLocation;
        public static string initialFont;
        public static WdColor initialColor;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
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
