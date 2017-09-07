<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmMain
    Inherits DevExpress.XtraEditors.XtraForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim SplashScreenManager1 As DevExpress.XtraSplashScreen.SplashScreenManager = New DevExpress.XtraSplashScreen.SplashScreenManager(Me, GetType(Global.PrjSWSTAP.SplashScreen1), True, True)
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMain))
        Me.PanelControl2 = New DevExpress.XtraEditors.PanelControl()
        Me.NavBarControl1 = New DevExpress.XtraNavBar.NavBarControl()
        Me.NavBarGroup1 = New DevExpress.XtraNavBar.NavBarGroup()
        Me.NavBarItem1 = New DevExpress.XtraNavBar.NavBarItem()
        Me.PanelControl3 = New DevExpress.XtraEditors.PanelControl()
        Me.LabelControl8 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl6 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.PanelControl4 = New DevExpress.XtraEditors.PanelControl()
        Me.BunifuFlatButton1 = New ns1.BunifuFlatButton()
        Me.BunifuFlatButton3 = New ns1.BunifuFlatButton()
        Me.LabelControl7 = New DevExpress.XtraEditors.LabelControl()
        Me.DefaultLookAndFeel1 = New DevExpress.LookAndFeel.DefaultLookAndFeel(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.XtraTabbedMdiManager1 = New DevExpress.XtraTabbedMdi.XtraTabbedMdiManager(Me.components)
        Me.ImgGrp = New DevExpress.Utils.ImageCollection(Me.components)
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl2.SuspendLayout()
        CType(Me.NavBarControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PanelControl3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl3.SuspendLayout()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        CType(Me.PanelControl4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl4.SuspendLayout()
        CType(Me.XtraTabbedMdiManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ImgGrp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplashScreenManager1
        '
        SplashScreenManager1.ClosingDelay = 500
        '
        'PanelControl2
        '
        Me.PanelControl2.Appearance.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.PanelControl2.Appearance.Options.UseBackColor = True
        Me.PanelControl2.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl2.Controls.Add(Me.NavBarControl1)
        Me.PanelControl2.Controls.Add(Me.PanelControl3)
        Me.PanelControl2.Dock = System.Windows.Forms.DockStyle.Left
        Me.PanelControl2.Location = New System.Drawing.Point(0, 0)
        Me.PanelControl2.Name = "PanelControl2"
        Me.PanelControl2.Size = New System.Drawing.Size(248, 485)
        Me.PanelControl2.TabIndex = 2
        '
        'NavBarControl1
        '
        Me.NavBarControl1.ActiveGroup = Me.NavBarGroup1
        Me.NavBarControl1.Appearance.GroupHeader.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NavBarControl1.Appearance.GroupHeader.ForeColor = System.Drawing.Color.DarkGreen
        Me.NavBarControl1.Appearance.GroupHeader.Options.UseFont = True
        Me.NavBarControl1.Appearance.GroupHeader.Options.UseForeColor = True
        Me.NavBarControl1.Appearance.Item.BorderColor = System.Drawing.Color.OliveDrab
        Me.NavBarControl1.Appearance.Item.Options.UseBorderColor = True
        Me.NavBarControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Flat
        Me.NavBarControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.NavBarControl1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.NavBarControl1.Groups.AddRange(New DevExpress.XtraNavBar.NavBarGroup() {Me.NavBarGroup1})
        Me.NavBarControl1.Items.AddRange(New DevExpress.XtraNavBar.NavBarItem() {Me.NavBarItem1})
        Me.NavBarControl1.Location = New System.Drawing.Point(0, 137)
        Me.NavBarControl1.LookAndFeel.SkinName = "Office 2016 Black"
        Me.NavBarControl1.Name = "NavBarControl1"
        Me.NavBarControl1.NavigationPaneGroupClientHeight = 500
        Me.NavBarControl1.OptionsNavPane.ExpandedWidth = 248
        Me.NavBarControl1.Size = New System.Drawing.Size(248, 348)
        Me.NavBarControl1.TabIndex = 6
        Me.NavBarControl1.Text = "NavBarControl1"
        Me.NavBarControl1.View = New DevExpress.XtraNavBar.ViewInfo.StandardSkinExplorerBarViewInfoRegistrator("Visual Studio 2013 Dark")
        '
        'NavBarGroup1
        '
        Me.NavBarGroup1.Caption = "HOME"
        Me.NavBarGroup1.Expanded = True
        Me.NavBarGroup1.ItemLinks.AddRange(New DevExpress.XtraNavBar.NavBarItemLink() {New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem1)})
        Me.NavBarGroup1.Name = "NavBarGroup1"
        Me.NavBarGroup1.SmallImage = Global.PrjSWSTAP.My.Resources.Resources.Home_16px
        '
        'NavBarItem1
        '
        Me.NavBarItem1.Caption = "EXIT"
        Me.NavBarItem1.Name = "NavBarItem1"
        '
        'PanelControl3
        '
        Me.PanelControl3.Appearance.BackColor = System.Drawing.Color.SteelBlue
        Me.PanelControl3.Appearance.Options.UseBackColor = True
        Me.PanelControl3.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl3.Controls.Add(Me.LabelControl8)
        Me.PanelControl3.Controls.Add(Me.LabelControl6)
        Me.PanelControl3.Controls.Add(Me.LabelControl5)
        Me.PanelControl3.Controls.Add(Me.LabelControl4)
        Me.PanelControl3.Controls.Add(Me.LabelControl3)
        Me.PanelControl3.Controls.Add(Me.LabelControl2)
        Me.PanelControl3.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelControl3.Location = New System.Drawing.Point(0, 0)
        Me.PanelControl3.Name = "PanelControl3"
        Me.PanelControl3.Size = New System.Drawing.Size(248, 137)
        Me.PanelControl3.TabIndex = 5
        '
        'LabelControl8
        '
        Me.LabelControl8.Appearance.Font = New System.Drawing.Font("Tahoma", 9.25!, System.Drawing.FontStyle.Bold)
        Me.LabelControl8.Appearance.Options.UseFont = True
        Me.LabelControl8.Location = New System.Drawing.Point(12, 108)
        Me.LabelControl8.Name = "LabelControl8"
        Me.LabelControl8.Size = New System.Drawing.Size(13, 14)
        Me.LabelControl8.TabIndex = 5
        Me.LabelControl8.Text = "IP"
        '
        'LabelControl6
        '
        Me.LabelControl6.Appearance.Font = New System.Drawing.Font("Tahoma", 9.25!, System.Drawing.FontStyle.Bold)
        Me.LabelControl6.Appearance.Options.UseFont = True
        Me.LabelControl6.Location = New System.Drawing.Point(12, 88)
        Me.LabelControl6.Name = "LabelControl6"
        Me.LabelControl6.Size = New System.Drawing.Size(54, 14)
        Me.LabelControl6.TabIndex = 4
        Me.LabelControl6.Text = "VERSION"
        '
        'LabelControl5
        '
        Me.LabelControl5.Appearance.Font = New System.Drawing.Font("Tahoma", 9.25!, System.Drawing.FontStyle.Bold)
        Me.LabelControl5.Appearance.Options.UseFont = True
        Me.LabelControl5.Location = New System.Drawing.Point(12, 69)
        Me.LabelControl5.Name = "LabelControl5"
        Me.LabelControl5.Size = New System.Drawing.Size(62, 14)
        Me.LabelControl5.TabIndex = 3
        Me.LabelControl5.Text = "LOCATION"
        '
        'LabelControl4
        '
        Me.LabelControl4.Appearance.Font = New System.Drawing.Font("Tahoma", 9.25!, System.Drawing.FontStyle.Bold)
        Me.LabelControl4.Appearance.Options.UseFont = True
        Me.LabelControl4.Location = New System.Drawing.Point(12, 50)
        Me.LabelControl4.Name = "LabelControl4"
        Me.LabelControl4.Size = New System.Drawing.Size(74, 14)
        Me.LabelControl4.TabIndex = 2
        Me.LabelControl4.Text = "MILL PLANT"
        '
        'LabelControl3
        '
        Me.LabelControl3.Appearance.Font = New System.Drawing.Font("Tahoma", 9.25!, System.Drawing.FontStyle.Bold)
        Me.LabelControl3.Appearance.Options.UseFont = True
        Me.LabelControl3.Location = New System.Drawing.Point(12, 31)
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(66, 14)
        Me.LabelControl3.TabIndex = 1
        Me.LabelControl3.Text = "SITE NAME"
        '
        'LabelControl2
        '
        Me.LabelControl2.Appearance.Font = New System.Drawing.Font("Tahoma", 9.25!, System.Drawing.FontStyle.Bold)
        Me.LabelControl2.Appearance.Options.UseFont = True
        Me.LabelControl2.Location = New System.Drawing.Point(12, 12)
        Me.LabelControl2.Name = "LabelControl2"
        Me.LabelControl2.Size = New System.Drawing.Size(64, 14)
        Me.LabelControl2.TabIndex = 0
        Me.LabelControl2.Text = "SITE CODE"
        '
        'PanelControl1
        '
        Me.PanelControl1.Appearance.BackColor = System.Drawing.Color.SteelBlue
        Me.PanelControl1.Appearance.BackColor2 = System.Drawing.Color.Green
        Me.PanelControl1.Appearance.Options.UseBackColor = True
        Me.PanelControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl1.Controls.Add(Me.LabelControl1)
        Me.PanelControl1.Controls.Add(Me.PanelControl4)
        Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Top
        Me.PanelControl1.Location = New System.Drawing.Point(248, 0)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(714, 40)
        Me.PanelControl1.TabIndex = 3
        '
        'LabelControl1
        '
        Me.LabelControl1.Appearance.Font = New System.Drawing.Font("Tahoma", 11.0!, System.Drawing.FontStyle.Bold)
        Me.LabelControl1.Appearance.Options.UseFont = True
        Me.LabelControl1.Location = New System.Drawing.Point(6, 9)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(242, 18)
        Me.LabelControl1.TabIndex = 2
        Me.LabelControl1.Text = "Warehouse Management System"
        '
        'PanelControl4
        '
        Me.PanelControl4.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.PanelControl4.Appearance.Options.UseBackColor = True
        Me.PanelControl4.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl4.ContentImageAlignment = System.Drawing.ContentAlignment.TopCenter
        Me.PanelControl4.Controls.Add(Me.BunifuFlatButton1)
        Me.PanelControl4.Controls.Add(Me.BunifuFlatButton3)
        Me.PanelControl4.Controls.Add(Me.LabelControl7)
        Me.PanelControl4.Dock = System.Windows.Forms.DockStyle.Right
        Me.PanelControl4.Location = New System.Drawing.Point(290, 0)
        Me.PanelControl4.Name = "PanelControl4"
        Me.PanelControl4.Size = New System.Drawing.Size(424, 40)
        Me.PanelControl4.TabIndex = 1
        '
        'BunifuFlatButton1
        '
        Me.BunifuFlatButton1.Activecolor = System.Drawing.Color.SteelBlue
        Me.BunifuFlatButton1.BackColor = System.Drawing.Color.Transparent
        Me.BunifuFlatButton1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BunifuFlatButton1.BorderRadius = 0
        Me.BunifuFlatButton1.ButtonText = "Sign In"
        Me.BunifuFlatButton1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BunifuFlatButton1.DisabledColor = System.Drawing.Color.Gray
        Me.BunifuFlatButton1.Dock = System.Windows.Forms.DockStyle.Right
        Me.BunifuFlatButton1.Font = New System.Drawing.Font("TeamViewer12", 9.0!, System.Drawing.FontStyle.Bold)
        Me.BunifuFlatButton1.Iconcolor = System.Drawing.Color.Transparent
        Me.BunifuFlatButton1.Iconimage = Global.PrjSWSTAP.My.Resources.Resources.Admin_16px
        Me.BunifuFlatButton1.Iconimage_right = Nothing
        Me.BunifuFlatButton1.Iconimage_right_Selected = Nothing
        Me.BunifuFlatButton1.Iconimage_Selected = Nothing
        Me.BunifuFlatButton1.IconMarginLeft = 0
        Me.BunifuFlatButton1.IconMarginRight = 0
        Me.BunifuFlatButton1.IconRightVisible = True
        Me.BunifuFlatButton1.IconRightZoom = 0R
        Me.BunifuFlatButton1.IconVisible = True
        Me.BunifuFlatButton1.IconZoom = 40.0R
        Me.BunifuFlatButton1.IsTab = False
        Me.BunifuFlatButton1.Location = New System.Drawing.Point(222, 0)
        Me.BunifuFlatButton1.Name = "BunifuFlatButton1"
        Me.BunifuFlatButton1.Normalcolor = System.Drawing.Color.Transparent
        Me.BunifuFlatButton1.OnHovercolor = System.Drawing.SystemColors.ActiveCaption
        Me.BunifuFlatButton1.OnHoverTextColor = System.Drawing.Color.White
        Me.BunifuFlatButton1.selected = False
        Me.BunifuFlatButton1.Size = New System.Drawing.Size(101, 40)
        Me.BunifuFlatButton1.TabIndex = 5
        Me.BunifuFlatButton1.Text = "Sign In"
        Me.BunifuFlatButton1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.BunifuFlatButton1.Textcolor = System.Drawing.Color.White
        Me.BunifuFlatButton1.TextFont = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'BunifuFlatButton3
        '
        Me.BunifuFlatButton3.Activecolor = System.Drawing.Color.Orange
        Me.BunifuFlatButton3.BackColor = System.Drawing.Color.Transparent
        Me.BunifuFlatButton3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BunifuFlatButton3.BorderRadius = 0
        Me.BunifuFlatButton3.ButtonText = "Config"
        Me.BunifuFlatButton3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BunifuFlatButton3.DisabledColor = System.Drawing.Color.Gray
        Me.BunifuFlatButton3.Dock = System.Windows.Forms.DockStyle.Right
        Me.BunifuFlatButton3.Font = New System.Drawing.Font("TeamViewer12", 9.0!, System.Drawing.FontStyle.Bold)
        Me.BunifuFlatButton3.Iconcolor = System.Drawing.Color.Transparent
        Me.BunifuFlatButton3.Iconimage = Global.PrjSWSTAP.My.Resources.Resources.SettingsW_16px
        Me.BunifuFlatButton3.Iconimage_right = Nothing
        Me.BunifuFlatButton3.Iconimage_right_Selected = Nothing
        Me.BunifuFlatButton3.Iconimage_Selected = Nothing
        Me.BunifuFlatButton3.IconMarginLeft = 0
        Me.BunifuFlatButton3.IconMarginRight = 0
        Me.BunifuFlatButton3.IconRightVisible = True
        Me.BunifuFlatButton3.IconRightZoom = 0R
        Me.BunifuFlatButton3.IconVisible = True
        Me.BunifuFlatButton3.IconZoom = 40.0R
        Me.BunifuFlatButton3.IsTab = False
        Me.BunifuFlatButton3.Location = New System.Drawing.Point(323, 0)
        Me.BunifuFlatButton3.Name = "BunifuFlatButton3"
        Me.BunifuFlatButton3.Normalcolor = System.Drawing.Color.Transparent
        Me.BunifuFlatButton3.OnHovercolor = System.Drawing.SystemColors.ActiveCaption
        Me.BunifuFlatButton3.OnHoverTextColor = System.Drawing.Color.White
        Me.BunifuFlatButton3.selected = False
        Me.BunifuFlatButton3.Size = New System.Drawing.Size(101, 40)
        Me.BunifuFlatButton3.TabIndex = 4
        Me.BunifuFlatButton3.Text = "Config"
        Me.BunifuFlatButton3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.BunifuFlatButton3.Textcolor = System.Drawing.Color.White
        Me.BunifuFlatButton3.TextFont = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'LabelControl7
        '
        Me.LabelControl7.Appearance.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.LabelControl7.Appearance.Options.UseFont = True
        Me.LabelControl7.Location = New System.Drawing.Point(163, 12)
        Me.LabelControl7.Name = "LabelControl7"
        Me.LabelControl7.Size = New System.Drawing.Size(55, 13)
        Me.LabelControl7.TabIndex = 3
        Me.LabelControl7.Text = "Welcome,"
        '
        'DefaultLookAndFeel1
        '
        Me.DefaultLookAndFeel1.LookAndFeel.SkinName = "Office 2016 Colorful"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'XtraTabbedMdiManager1
        '
        Me.XtraTabbedMdiManager1.Appearance.BackColor = System.Drawing.Color.DarkGray
        Me.XtraTabbedMdiManager1.Appearance.BackColor2 = System.Drawing.Color.DarkGray
        Me.XtraTabbedMdiManager1.Appearance.BorderColor = System.Drawing.Color.DarkGray
        Me.XtraTabbedMdiManager1.Appearance.Options.UseBackColor = True
        Me.XtraTabbedMdiManager1.Appearance.Options.UseBorderColor = True
        Me.XtraTabbedMdiManager1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.XtraTabbedMdiManager1.BorderStylePage = DevExpress.XtraEditors.Controls.BorderStyles.Office2003
        Me.XtraTabbedMdiManager1.HeaderOrientation = DevExpress.XtraTab.TabOrientation.Horizontal
        Me.XtraTabbedMdiManager1.MdiParent = Me
        Me.XtraTabbedMdiManager1.ShowHeaderFocus = DevExpress.Utils.DefaultBoolean.[True]
        '
        'ImgGrp
        '
        Me.ImgGrp.ImageStream = CType(resources.GetObject("ImgGrp.ImageStream"), DevExpress.Utils.ImageCollectionStreamer)
        Me.ImgGrp.Images.SetKeyName(0, "About16.png")
        Me.ImgGrp.Images.SetKeyName(1, "AddColumn16.png")
        Me.ImgGrp.Images.SetKeyName(2, "Barcode16.png")
        Me.ImgGrp.Images.SetKeyName(3, "BarcodeScanner16.png")
        Me.ImgGrp.Images.SetKeyName(4, "Cancel16.png")
        Me.ImgGrp.Images.SetKeyName(5, "CloseWindow16.png")
        Me.ImgGrp.Images.SetKeyName(6, "Contacts16.png")
        Me.ImgGrp.Images.SetKeyName(7, "Dashboard16.png")
        Me.ImgGrp.Images.SetKeyName(8, "Dashboard216.png")
        Me.ImgGrp.Images.SetKeyName(9, "DataBackup16.png")
        Me.ImgGrp.Images.SetKeyName(10, "DataConfiguration16.png")
        Me.ImgGrp.Images.SetKeyName(11, "DeleteColumn16.png")
        Me.ImgGrp.Images.SetKeyName(12, "Excel16.png")
        Me.ImgGrp.Images.SetKeyName(13, "Exit16.png")
        Me.ImgGrp.Images.SetKeyName(14, "ExitSign16.png")
        Me.ImgGrp.Images.SetKeyName(15, "Factory16.png")
        Me.ImgGrp.Images.SetKeyName(16, "ForkLift16.png")
        Me.ImgGrp.Images.SetKeyName(17, "Home16.png")
        Me.ImgGrp.Images.SetKeyName(18, "Import16px.png")
        Me.ImgGrp.Images.SetKeyName(19, "Info16.png")
        Me.ImgGrp.Images.SetKeyName(20, "KartuStock16.png")
        Me.ImgGrp.Images.SetKeyName(21, "MaximizeWindow16.png")
        Me.ImgGrp.Images.SetKeyName(22, "MinimizeWindow16.png")
        Me.ImgGrp.Images.SetKeyName(23, "MoveStock16.png")
        Me.ImgGrp.Images.SetKeyName(24, "NewProduct16.png")
        Me.ImgGrp.Images.SetKeyName(25, "Ok16.png")
        Me.ImgGrp.Images.SetKeyName(26, "PriceTag16.png")
        Me.ImgGrp.Images.SetKeyName(27, "Rack16.png")
        Me.ImgGrp.Images.SetKeyName(28, "Report16.png")
        Me.ImgGrp.Images.SetKeyName(29, "Save16.png")
        Me.ImgGrp.Images.SetKeyName(30, "Settings16.png")
        Me.ImgGrp.Images.SetKeyName(31, "ShoppingCartLoaded16.png")
        Me.ImgGrp.Images.SetKeyName(32, "Sync16.png")
        Me.ImgGrp.Images.SetKeyName(33, "SystemAdmin16.png")
        Me.ImgGrp.Images.SetKeyName(34, "Tools16.png")
        Me.ImgGrp.Images.SetKeyName(35, "Transaksi16.png")
        Me.ImgGrp.Images.SetKeyName(36, "UserGroup16.png")
        Me.ImgGrp.Images.SetKeyName(37, "Work16.png")
        '
        'FrmMain
        '
        Me.Appearance.BackColor = System.Drawing.Color.White
        Me.Appearance.ForeColor = System.Drawing.Color.White
        Me.Appearance.Options.UseBackColor = True
        Me.Appearance.Options.UseForeColor = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImageLayoutStore = System.Windows.Forms.ImageLayout.Tile
        Me.BackgroundImageStore = CType(resources.GetObject("$this.BackgroundImageStore"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(962, 485)
        Me.Controls.Add(Me.PanelControl1)
        Me.Controls.Add(Me.PanelControl2)
        Me.DoubleBuffered = True
        Me.FormBorderEffect = DevExpress.XtraEditors.FormBorderEffect.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.LookAndFeel.SkinName = "Office 2013"
        Me.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.UltraFlat
        Me.Name = "FrmMain"
        Me.Text = "PT AIMFOOD MANUFACTURING INDONESIA"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl2.ResumeLayout(False)
        CType(Me.NavBarControl1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PanelControl3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl3.ResumeLayout(False)
        Me.PanelControl3.PerformLayout()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.PanelControl1.PerformLayout()
        CType(Me.PanelControl4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl4.ResumeLayout(False)
        Me.PanelControl4.PerformLayout()
        CType(Me.XtraTabbedMdiManager1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ImgGrp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PanelControl2 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents PanelControl3 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents DefaultLookAndFeel1 As DevExpress.LookAndFeel.DefaultLookAndFeel
    Friend WithEvents PanelControl4 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents NavBarControl1 As DevExpress.XtraNavBar.NavBarControl
    Friend WithEvents NavBarGroup1 As DevExpress.XtraNavBar.NavBarGroup
    Friend WithEvents NavBarItem1 As DevExpress.XtraNavBar.NavBarItem
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl6 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents LabelControl7 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents Timer1 As Timer
    Friend WithEvents XtraTabbedMdiManager1 As DevExpress.XtraTabbedMdi.XtraTabbedMdiManager
    Friend WithEvents BunifuFlatButton3 As ns1.BunifuFlatButton
    Friend WithEvents BunifuFlatButton1 As ns1.BunifuFlatButton
    Friend WithEvents ImgGrp As DevExpress.Utils.ImageCollection
    Friend WithEvents LabelControl8 As DevExpress.XtraEditors.LabelControl
End Class
