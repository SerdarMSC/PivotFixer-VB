<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        btnSelect = New Button()
        btnStart = New Button()
        lblStatus = New Label()
        progressBar1 = New ProgressBar()
        lblResult = New Label()
        SuspendLayout()
        ' 
        ' btnSelect
        ' 
        btnSelect.Location = New Point(12, 12)
        btnSelect.Name = "btnSelect"
        btnSelect.Size = New Size(198, 39)
        btnSelect.TabIndex = 0
        btnSelect.Text = "Dosya Sec"
        btnSelect.UseVisualStyleBackColor = True
        ' 
        ' btnStart
        ' 
        btnStart.Location = New Point(216, 12)
        btnStart.Name = "btnStart"
        btnStart.Size = New Size(71, 39)
        btnStart.TabIndex = 1
        btnStart.Text = "Basla"
        btnStart.UseVisualStyleBackColor = True
        ' 
        ' lblStatus
        ' 
        lblStatus.AutoSize = True
        lblStatus.Location = New Point(12, 67)
        lblStatus.Name = "lblStatus"
        lblStatus.Size = New Size(43, 15)
        lblStatus.TabIndex = 2
        lblStatus.Text = "Hazir..."
        ' 
        ' progressBar1
        ' 
        progressBar1.Location = New Point(12, 107)
        progressBar1.Name = "progressBar1"
        progressBar1.Size = New Size(275, 23)
        progressBar1.TabIndex = 3
        ' 
        ' lblResult
        ' 
        lblResult.AutoSize = True
        lblResult.Location = New Point(14, 89)
        lblResult.Name = "lblResult"
        lblResult.Size = New Size(41, 15)
        lblResult.TabIndex = 4
        lblResult.Text = "Label1"
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(298, 145)
        Controls.Add(lblResult)
        Controls.Add(progressBar1)
        Controls.Add(lblStatus)
        Controls.Add(btnStart)
        Controls.Add(btnSelect)
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        MaximizeBox = False
        MdiChildrenMinimizedAnchorBottom = False
        MinimizeBox = False
        Name = "Form1"
        Text = "Pivot Data Duzenleyici"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents btnSelect As Button
    Friend WithEvents btnStart As Button
    Friend WithEvents lblStatus As Label
    Friend WithEvents progressBar1 As ProgressBar
    Friend WithEvents lblResult As Label

End Class
