namespace WordAddIn1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_AddBookMark = this.Factory.CreateRibbonButton();
            this.btn_Location = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btn_Ident = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "aaaa";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_AddBookMark);
            this.group1.Items.Add(this.btn_Location);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btn_AddBookMark
            // 
            this.btn_AddBookMark.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_AddBookMark.Label = "添加书签";
            this.btn_AddBookMark.Name = "btn_AddBookMark";
            this.btn_AddBookMark.ShowImage = true;
            this.btn_AddBookMark.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_AddBookMark_Click);
            // 
            // btn_Location
            // 
            this.btn_Location.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Location.Label = "定位";
            this.btn_Location.Name = "btn_Location";
            this.btn_Location.ShowImage = true;
            this.btn_Location.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Location_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btn_Ident);
            this.group2.Label = "group2";
            this.group2.Name = "group2";
            // 
            // btn_Ident
            // 
            this.btn_Ident.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Ident.Label = "button1";
            this.btn_Ident.Name = "btn_Ident";
            this.btn_Ident.ShowImage = true;
            this.btn_Ident.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Ident_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_AddBookMark;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Location;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Ident;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
