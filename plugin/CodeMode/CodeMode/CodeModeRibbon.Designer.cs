
namespace CodeMode
{
    partial class CodeModeRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CodeModeRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CodeModeRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.codeMode = this.Factory.CreateRibbonGroup();
            this.toggleCodeMode = this.Factory.CreateRibbonToggleButton();
            this.selectCodeColor = this.Factory.CreateRibbonButton();
            this.selectBackgroundColor = this.Factory.CreateRibbonButton();
            this.codeFont = this.Factory.CreateRibbonComboBox();
            this.tab1.SuspendLayout();
            this.codeMode.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.codeMode);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // codeMode
            // 
            this.codeMode.Items.Add(this.toggleCodeMode);
            this.codeMode.Items.Add(this.selectCodeColor);
            this.codeMode.Items.Add(this.selectBackgroundColor);
            this.codeMode.Items.Add(this.codeFont);
            this.codeMode.Label = "Code Mode";
            this.codeMode.Name = "codeMode";
            // 
            // toggleCodeMode
            // 
            this.toggleCodeMode.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleCodeMode.Description = "Toggles Code Mode";
            this.toggleCodeMode.Image = ((System.Drawing.Image)(resources.GetObject("toggleCodeMode.Image")));
            this.toggleCodeMode.Label = "Toggle Code Mode";
            this.toggleCodeMode.Name = "toggleCodeMode";
            this.toggleCodeMode.ShowImage = true;
            this.toggleCodeMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleCodeMode_Click);
            // 
            // selectCodeColor
            // 
            this.selectCodeColor.Label = "Select Code Colour";
            this.selectCodeColor.Name = "selectCodeColor";
            this.selectCodeColor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectCodeColor_Click);
            // 
            // selectBackgroundColor
            // 
            this.selectBackgroundColor.Label = "Select Background Colour";
            this.selectBackgroundColor.Name = "selectBackgroundColor";
            this.selectBackgroundColor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectBackgroundColor_Click);
            // 
            // codeFont
            // 
            this.codeFont.Label = "Code Font";
            this.codeFont.MaxLength = 20;
            this.codeFont.Name = "codeFont";
            this.codeFont.Text = null;
            this.codeFont.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.codeFont_TextChanged);
            // 
            // CodeModeRibbon
            // 
            this.Name = "CodeModeRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CodeModeRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.codeMode.ResumeLayout(false);
            this.codeMode.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup codeMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleCodeMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectCodeColor;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectBackgroundColor;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox codeFont;
    }

    partial class ThisRibbonCollection
    {
        internal CodeModeRibbon CodeModeRibbon
        {
            get { return this.GetRibbon<CodeModeRibbon>(); }
        }
    }
}
