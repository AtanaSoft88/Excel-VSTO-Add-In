
namespace excelAddInTest
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.loadFromDbBtn = this.Factory.CreateRibbonButton();
            this.calcCommissionBtn = this.Factory.CreateRibbonButton();
            this.clearBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.loadFromDbBtn);
            this.group1.Items.Add(this.calcCommissionBtn);
            this.group1.Items.Add(this.clearBtn);
            this.group1.Label = "*";
            this.group1.Name = "group1";
            // 
            // loadFromDbBtn
            // 
            this.loadFromDbBtn.Label = "Load from DB";
            this.loadFromDbBtn.Name = "loadFromDbBtn";
            this.loadFromDbBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.loadFromDbBtn_Click);
            // 
            // calcCommissionBtn
            // 
            this.calcCommissionBtn.Label = "Calculate Commision";
            this.calcCommissionBtn.Name = "calcCommissionBtn";
            this.calcCommissionBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.calcCommissionBtn_Click);
            // 
            // clearBtn
            // 
            this.clearBtn.Label = "Clear Commission";
            this.clearBtn.Name = "clearBtn";
            this.clearBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearBtn_Click);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton calcCommissionBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton clearBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loadFromDbBtn;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon Ribbon1
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
