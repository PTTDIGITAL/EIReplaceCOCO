<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

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
        Me.components = New System.ComponentModel.Container()
        Me.pbExit = New System.Windows.Forms.PictureBox()
        Me.pbExport = New System.Windows.Forms.PictureBox()
        Me.pbImport = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblDateTime = New System.Windows.Forms.Label()
        Me.txtTransLog = New System.Windows.Forms.TextBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        CType(Me.pbExit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbExport, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbImport, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'pbExit
        '
        Me.pbExit.BackColor = System.Drawing.Color.Transparent
        Me.pbExit.BackgroundImage = Global.EIReplaceCOCO.My.Resources.Resources.Close
        Me.pbExit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.pbExit.Cursor = System.Windows.Forms.Cursors.Hand
        Me.pbExit.Location = New System.Drawing.Point(500, 105)
        Me.pbExit.Name = "pbExit"
        Me.pbExit.Size = New System.Drawing.Size(116, 45)
        Me.pbExit.TabIndex = 0
        Me.pbExit.TabStop = False
        '
        'pbExport
        '
        Me.pbExport.BackColor = System.Drawing.Color.Transparent
        Me.pbExport.BackgroundImage = Global.EIReplaceCOCO.My.Resources.Resources.export_th__2_
        Me.pbExport.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.pbExport.Cursor = System.Windows.Forms.Cursors.Hand
        Me.pbExport.Location = New System.Drawing.Point(12, 105)
        Me.pbExport.Name = "pbExport"
        Me.pbExport.Size = New System.Drawing.Size(120, 45)
        Me.pbExport.TabIndex = 1
        Me.pbExport.TabStop = False
        '
        'pbImport
        '
        Me.pbImport.BackColor = System.Drawing.Color.Transparent
        Me.pbImport.BackgroundImage = Global.EIReplaceCOCO.My.Resources.Resources.Import_th
        Me.pbImport.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.pbImport.Cursor = System.Windows.Forms.Cursors.Hand
        Me.pbImport.Location = New System.Drawing.Point(138, 105)
        Me.pbImport.Name = "pbImport"
        Me.pbImport.Size = New System.Drawing.Size(121, 45)
        Me.pbImport.TabIndex = 2
        Me.pbImport.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.BackgroundImage = Global.EIReplaceCOCO.My.Resources.Resources.imgpttgdt
        Me.Panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel1.Location = New System.Drawing.Point(0, -1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(227, 100)
        Me.Panel1.TabIndex = 3
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.Controls.Add(Me.lblDateTime)
        Me.Panel2.Location = New System.Drawing.Point(227, -1)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(415, 100)
        Me.Panel2.TabIndex = 5
        '
        'lblDateTime
        '
        Me.lblDateTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblDateTime.Location = New System.Drawing.Point(187, 0)
        Me.lblDateTime.Name = "lblDateTime"
        Me.lblDateTime.Size = New System.Drawing.Size(218, 42)
        Me.lblDateTime.TabIndex = 0
        Me.lblDateTime.Text = "Label1"
        '
        'txtTransLog
        '
        Me.txtTransLog.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtTransLog.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.txtTransLog.Location = New System.Drawing.Point(10, 177)
        Me.txtTransLog.Multiline = True
        Me.txtTransLog.Name = "txtTransLog"
        Me.txtTransLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtTransLog.Size = New System.Drawing.Size(605, 275)
        Me.txtTransLog.TabIndex = 7
        '
        'Timer1
        '
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.ClientSize = New System.Drawing.Size(641, 500)
        Me.ControlBox = False
        Me.Controls.Add(Me.txtTransLog)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.pbImport)
        Me.Controls.Add(Me.pbExport)
        Me.Controls.Add(Me.pbExit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "EI REPLACE COCO V1.0"
        CType(Me.pbExit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbExport, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbImport, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents pbExit As PictureBox
    Friend WithEvents pbExport As PictureBox
    Friend WithEvents pbImport As PictureBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents txtTransLog As TextBox
    Friend WithEvents Timer1 As Timer
    Friend WithEvents lblDateTime As Label
End Class
