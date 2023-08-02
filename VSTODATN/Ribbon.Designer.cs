namespace VSTODATN
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Algorithm = this.Factory.CreateRibbonGroup();
            this.ImageCells = this.Factory.CreateRibbonButton();
            this.Speech = this.Factory.CreateRibbonGroup();
            this.SpeechText = this.Factory.CreateRibbonButton();
            this.Finance = this.Factory.CreateRibbonGroup();
            this.Exchange = this.Factory.CreateRibbonButton();
            this.ReadNumber = this.Factory.CreateRibbonButton();
            this.Json = this.Factory.CreateRibbonGroup();
            this.ExportJson = this.Factory.CreateRibbonButton();
            this.ImportJson = this.Factory.CreateRibbonButton();
            this.ExportJsonFormat = this.Factory.CreateRibbonButton();
            this.Instructions = this.Factory.CreateRibbonGroup();
            this.Instruction = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Algorithm.SuspendLayout();
            this.Speech.SuspendLayout();
            this.Finance.SuspendLayout();
            this.Json.SuspendLayout();
            this.Instructions.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Algorithm);
            this.tab1.Groups.Add(this.Speech);
            this.tab1.Groups.Add(this.Finance);
            this.tab1.Groups.Add(this.Json);
            this.tab1.Groups.Add(this.Instructions);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // Algorithm
            // 
            this.Algorithm.Items.Add(this.ImageCells);
            this.Algorithm.Label = "Algorithm";
            this.Algorithm.Name = "Algorithm";
            // 
            // ImageCells
            // 
            this.ImageCells.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ImageCells.Image = global::VSTODATN.Properties.Resources.picture;
            this.ImageCells.Label = "ImageCells";
            this.ImageCells.Name = "ImageCells";
            this.ImageCells.ShowImage = true;
            this.ImageCells.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonImage2Cells_Click);
            // 
            // Speech
            // 
            this.Speech.Items.Add(this.SpeechText);
            this.Speech.Label = "Speech";
            this.Speech.Name = "Speech";
            // 
            // SpeechText
            // 
            this.SpeechText.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SpeechText.Image = global::VSTODATN.Properties.Resources.file;
            this.SpeechText.Label = "SpeechText(Eng)";
            this.SpeechText.Name = "SpeechText";
            this.SpeechText.ShowImage = true;
            this.SpeechText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SpeedText_Click);
            // 
            // Finance
            // 
            this.Finance.Items.Add(this.Exchange);
            this.Finance.Items.Add(this.ReadNumber);
            this.Finance.Label = "Finance";
            this.Finance.Name = "Finance";
            // 
            // Exchange
            // 
            this.Exchange.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Exchange.Image = global::VSTODATN.Properties.Resources.exchange;
            this.Exchange.Label = "Exchange";
            this.Exchange.Name = "Exchange";
            this.Exchange.ShowImage = true;
            this.Exchange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Exchange_Click);
            // 
            // ReadNumber
            // 
            this.ReadNumber.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ReadNumber.Image = global::VSTODATN.Properties.Resources.number_blocks;
            this.ReadNumber.Label = "ReadNumber";
            this.ReadNumber.Name = "ReadNumber";
            this.ReadNumber.ShowImage = true;
            this.ReadNumber.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReadNumber_Click);
            // 
            // Json
            // 
            this.Json.Items.Add(this.ExportJson);
            this.Json.Items.Add(this.ImportJson);
            this.Json.Items.Add(this.ExportJsonFormat);
            this.Json.Label = "Json";
            this.Json.Name = "Json";
            // 
            // ExportJson
            // 
            this.ExportJson.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ExportJson.Image = global::VSTODATN.Properties.Resources.download__1_;
            this.ExportJson.Label = "ExportJson";
            this.ExportJson.Name = "ExportJson";
            this.ExportJson.ShowImage = true;
            this.ExportJson.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportJson_Click);
            // 
            // ImportJson
            // 
            this.ImportJson.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ImportJson.Image = global::VSTODATN.Properties.Resources.upload;
            this.ImportJson.Label = "ImportJson";
            this.ImportJson.Name = "ImportJson";
            this.ImportJson.ShowImage = true;
            this.ImportJson.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportJson_Click);
            // 
            // ExportJsonFormat
            // 
            this.ExportJsonFormat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ExportJsonFormat.Image = global::VSTODATN.Properties.Resources.download__1_;
            this.ExportJsonFormat.Label = "ExportJsonFormat";
            this.ExportJsonFormat.Name = "ExportJsonFormat";
            this.ExportJsonFormat.ShowImage = true;
            this.ExportJsonFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportJsonFormat_Click);
            // 
            // Instructions
            // 
            this.Instructions.Items.Add(this.Instruction);
            this.Instructions.Label = "Instructions";
            this.Instructions.Name = "Instructions";
            // 
            // Instruction
            // 
            this.Instruction.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Instruction.Image = global::VSTODATN.Properties.Resources.electric_bicycle;
            this.Instruction.Label = "Instruction";
            this.Instruction.Name = "Instruction";
            this.Instruction.ShowImage = true;
            this.Instruction.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Instruction_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Algorithm.ResumeLayout(false);
            this.Algorithm.PerformLayout();
            this.Speech.ResumeLayout(false);
            this.Speech.PerformLayout();
            this.Finance.ResumeLayout(false);
            this.Finance.PerformLayout();
            this.Json.ResumeLayout(false);
            this.Json.PerformLayout();
            this.Instructions.ResumeLayout(false);
            this.Instructions.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Algorithm;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Speech;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Finance;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Json;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Instructions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ImageCells;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SpeechText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Exchange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ReadNumber;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportJson;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ImportJson;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportJsonFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Instruction;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
