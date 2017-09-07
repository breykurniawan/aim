<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SplashScreen1
    Inherits DevExpress.XtraSplashScreen.SplashScreen

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.labelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.labelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.marqueeProgressBarControl1 = New DevExpress.XtraEditors.MarqueeProgressBarControl()
        Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
        Me.pictureEdit2 = New DevExpress.XtraEditors.PictureEdit()
        CType(Me.marqueeProgressBarControl1.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureEdit2.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'labelControl2
        '
        Me.labelControl2.Location = New System.Drawing.Point(23, 164)
        Me.labelControl2.Name = "labelControl2"
        Me.labelControl2.Size = New System.Drawing.Size(50, 13)
        Me.labelControl2.TabIndex = 12
        Me.labelControl2.Text = "Starting..."
        '
        'labelControl1
        '
        Me.labelControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.labelControl1.Location = New System.Drawing.Point(23, 201)
        Me.labelControl1.Name = "labelControl1"
        Me.labelControl1.Size = New System.Drawing.Size(115, 13)
        Me.labelControl1.TabIndex = 11
        Me.labelControl1.Text = "Copyright © 1998-2013"
        '
        'marqueeProgressBarControl1
        '
        Me.marqueeProgressBarControl1.EditValue = 0
        Me.marqueeProgressBarControl1.Location = New System.Drawing.Point(23, 183)
        Me.marqueeProgressBarControl1.Name = "marqueeProgressBarControl1"
        Me.marqueeProgressBarControl1.Properties.EndColor = System.Drawing.Color.LimeGreen
        Me.marqueeProgressBarControl1.Properties.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
        Me.marqueeProgressBarControl1.Properties.StartColor = System.Drawing.Color.LimeGreen
        Me.marqueeProgressBarControl1.Size = New System.Drawing.Size(404, 12)
        Me.marqueeProgressBarControl1.TabIndex = 10
        '
        'LabelControl3
        '
        Me.LabelControl3.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.LabelControl3.Location = New System.Drawing.Point(23, 220)
        Me.LabelControl3.Name = "LabelControl3"
        Me.LabelControl3.Size = New System.Drawing.Size(35, 13)
        Me.LabelControl3.TabIndex = 15
        Me.LabelControl3.Text = "Version"
        '
        'pictureEdit2
        '
        Me.pictureEdit2.Cursor = System.Windows.Forms.Cursors.Default
        Me.pictureEdit2.EditValue = Global.PrjSWSTAP.My.Resources.Resources.SPLFC
        Me.pictureEdit2.Location = New System.Drawing.Point(23, 12)
        Me.pictureEdit2.Name = "pictureEdit2"
        Me.pictureEdit2.Properties.AllowFocused = False
        Me.pictureEdit2.Properties.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.pictureEdit2.Properties.Appearance.Options.UseBackColor = True
        Me.pictureEdit2.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.pictureEdit2.Properties.ShowMenu = False
        Me.pictureEdit2.Properties.SizeMode = DevExpress.XtraEditors.Controls.PictureSizeMode.Zoom
        Me.pictureEdit2.Properties.ZoomAccelerationFactor = 1.0R
        Me.pictureEdit2.Size = New System.Drawing.Size(404, 146)
        Me.pictureEdit2.TabIndex = 14
        '
        'SplashScreen1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(450, 257)
        Me.Controls.Add(Me.LabelControl3)
        Me.Controls.Add(Me.pictureEdit2)
        Me.Controls.Add(Me.labelControl2)
        Me.Controls.Add(Me.labelControl1)
        Me.Controls.Add(Me.marqueeProgressBarControl1)
        Me.Name = "SplashScreen1"
        Me.Text = "Form1"
        CType(Me.marqueeProgressBarControl1.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureEdit2.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents labelControl2 As DevExpress.XtraEditors.LabelControl
    Private WithEvents labelControl1 As DevExpress.XtraEditors.LabelControl
    Private WithEvents marqueeProgressBarControl1 As DevExpress.XtraEditors.MarqueeProgressBarControl
    Private WithEvents pictureEdit2 As DevExpress.XtraEditors.PictureEdit
    Private WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
End Class
