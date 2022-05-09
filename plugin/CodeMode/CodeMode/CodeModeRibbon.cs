using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualBasic;

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

    public partial class CodeModeRibbon
    {
        ColorDialog colorPicker;
        static WdColor codeColor;
        static WdColor bgColor;

        private void CodeModeRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            foreach (FontFamily font in FontFamily.Families)
            {

                this.codeFont.Items.Add(new FontItem() { Label = font.Name, Tag = font.Name });
            }

        }

        private void toggleCodeMode_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.toggleCodeMode.Checked)
            {
                Selection initialLocation = Globals.ThisAddIn.Application.Selection;
                ThisAddIn.initialLocation = initialLocation.Range;

                string initialFont = initialLocation.Font.Name;
                ThisAddIn.initialFont = initialFont;

                initialLocation.Font.Name = "Courier New";


                Debug.WriteLine(initialLocation.Start);
            } else
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


        private void selectCodeColor_Click(object sender, RibbonControlEventArgs e)
        {
            colorPicker = new ColorDialog();
            colorPicker.AllowFullOpen = true;
            colorPicker.ShowHelp = true;
            colorPicker.ShowDialog();

            string colorString = colorPicker.Color.R.ToString() + colorPicker.Color.G.ToString() + colorPicker.Color.B.ToString();
            codeColor = (WdColor)Int32.Parse(colorString);

        }

        private void selectBackgroundColor_Click(object sender, RibbonControlEventArgs e)
        {
            colorPicker = new ColorDialog();
            colorPicker.AllowFullOpen = true;
            colorPicker.ShowHelp = true;
            colorPicker.ShowDialog();

            string colorString = colorPicker.Color.R.ToString() + colorPicker.Color.G.ToString() + colorPicker.Color.B.ToString();
            bgColor = (WdColor)Int32.Parse(colorString);
        }

        private void codeFont_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
