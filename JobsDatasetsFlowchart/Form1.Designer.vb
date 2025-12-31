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
    Label1 = New Label()
    txtSandboxFolder = New TextBox()
    Label2 = New Label()
    Label3 = New Label()
    txtOutputFolder = New TextBox()
    Label4 = New Label()
    txtModelFilename = New TextBox()
    btnLoadJobs = New Button()
    lbAvailableJobs = New ListBox()
    lbSelectedJobs = New ListBox()
    btnSelectJobItem = New Button()
    lblStatus = New Label()
    btnSelectJobItems = New Button()
    btnDeselectJobItem = New Button()
    btnDeselectJobItems = New Button()
    btnClearAvailableJobs = New Button()
    btnClearSelectedJobs = New Button()
    Label5 = New Label()
    txtProjectFilename = New TextBox()
    Label7 = New Label()
    btnGo = New Button()
    btnClose = New Button()
    Label8 = New Label()
    txtPumlFolder = New TextBox()
    cbLibrary = New CheckBox()
    cbPDS = New CheckBox()
    cbGDG = New CheckBox()
    cbFile = New CheckBox()
    Label9 = New Label()
    lblJobsLoaded = New Label()
    cbSQL = New CheckBox()
    SuspendLayout()
    ' 
    ' Label1
    ' 
    Label1.AutoSize = True
    Label1.Location = New Point(23, 19)
    Label1.Name = "Label1"
    Label1.Size = New Size(86, 25)
    Label1.TabIndex = 0
    Label1.Text = "Sandbox:"
    ' 
    ' txtSandboxFolder
    ' 
    txtSandboxFolder.Location = New Point(115, 16)
    txtSandboxFolder.Name = "txtSandboxFolder"
    txtSandboxFolder.Size = New Size(697, 31)
    txtSandboxFolder.TabIndex = 1
    ' 
    ' Label2
    ' 
    Label2.AutoSize = True
    Label2.Font = New Font("Segoe UI", 6F, FontStyle.Italic)
    Label2.Location = New Point(594, 56)
    Label2.Name = "Label2"
    Label2.Size = New Size(210, 15)
    Label2.TabIndex = 2
    Label2.Text = "Value set from last run of ADDILite.exe"
    ' 
    ' Label3
    ' 
    Label3.AutoSize = True
    Label3.Location = New Point(19, 71)
    Label3.Name = "Label3"
    Label3.Size = New Size(128, 25)
    Label3.TabIndex = 3
    Label3.Text = "Output Folder:"
    ' 
    ' txtOutputFolder
    ' 
    txtOutputFolder.Location = New Point(185, 68)
    txtOutputFolder.Name = "txtOutputFolder"
    txtOutputFolder.Size = New Size(150, 31)
    txtOutputFolder.TabIndex = 2
    txtOutputFolder.Text = "\OUTPUT"
    ' 
    ' Label4
    ' 
    Label4.AutoSize = True
    Label4.Location = New Point(19, 114)
    Label4.Name = "Label4"
    Label4.Size = New Size(156, 25)
    Label4.TabIndex = 5
    Label4.Text = "Model Worksheet:"
    ' 
    ' txtModelFilename
    ' 
    txtModelFilename.Location = New Point(185, 111)
    txtModelFilename.Name = "txtModelFilename"
    txtModelFilename.Size = New Size(233, 31)
    txtModelFilename.TabIndex = 3
    txtModelFilename.Text = "ADDILite.xlsx"
    ' 
    ' btnLoadJobs
    ' 
    btnLoadJobs.Location = New Point(500, 156)
    btnLoadJobs.Name = "btnLoadJobs"
    btnLoadJobs.Size = New Size(112, 34)
    btnLoadJobs.TabIndex = 8
    btnLoadJobs.Text = "Load Jobs"
    btnLoadJobs.UseVisualStyleBackColor = True
    ' 
    ' lbAvailableJobs
    ' 
    lbAvailableJobs.Enabled = False
    lbAvailableJobs.FormattingEnabled = True
    lbAvailableJobs.ItemHeight = 25
    lbAvailableJobs.Location = New Point(31, 245)
    lbAvailableJobs.Name = "lbAvailableJobs"
    lbAvailableJobs.SelectionMode = SelectionMode.MultiSimple
    lbAvailableJobs.Size = New Size(180, 204)
    lbAvailableJobs.TabIndex = 9
    ' 
    ' lbSelectedJobs
    ' 
    lbSelectedJobs.Enabled = False
    lbSelectedJobs.FormattingEnabled = True
    lbSelectedJobs.ItemHeight = 25
    lbSelectedJobs.Location = New Point(306, 245)
    lbSelectedJobs.Name = "lbSelectedJobs"
    lbSelectedJobs.SelectionMode = SelectionMode.MultiSimple
    lbSelectedJobs.Size = New Size(180, 204)
    lbSelectedJobs.TabIndex = 14
    ' 
    ' btnSelectJobItem
    ' 
    btnSelectJobItem.Enabled = False
    btnSelectJobItem.Font = New Font("Segoe UI", 9F, FontStyle.Bold)
    btnSelectJobItem.Location = New Point(229, 245)
    btnSelectJobItem.Name = "btnSelectJobItem"
    btnSelectJobItem.Size = New Size(51, 34)
    btnSelectJobItem.TabIndex = 10
    btnSelectJobItem.Text = ">"
    btnSelectJobItem.UseVisualStyleBackColor = True
    ' 
    ' lblStatus
    ' 
    lblStatus.AutoSize = True
    lblStatus.Location = New Point(23, 519)
    lblStatus.Name = "lblStatus"
    lblStatus.Size = New Size(222, 25)
    lblStatus.TabIndex = 12
    lblStatus.Text = "Click the Load Jobs button"
    ' 
    ' btnSelectJobItems
    ' 
    btnSelectJobItems.Enabled = False
    btnSelectJobItems.Location = New Point(229, 284)
    btnSelectJobItems.Name = "btnSelectJobItems"
    btnSelectJobItems.Size = New Size(51, 34)
    btnSelectJobItems.TabIndex = 11
    btnSelectJobItems.Text = ">>"
    btnSelectJobItems.UseVisualStyleBackColor = True
    ' 
    ' btnDeselectJobItem
    ' 
    btnDeselectJobItem.Enabled = False
    btnDeselectJobItem.Font = New Font("Segoe UI", 9F, FontStyle.Bold)
    btnDeselectJobItem.Location = New Point(229, 343)
    btnDeselectJobItem.Name = "btnDeselectJobItem"
    btnDeselectJobItem.Size = New Size(51, 34)
    btnDeselectJobItem.TabIndex = 12
    btnDeselectJobItem.Text = "<"
    btnDeselectJobItem.UseVisualStyleBackColor = True
    ' 
    ' btnDeselectJobItems
    ' 
    btnDeselectJobItems.Enabled = False
    btnDeselectJobItems.Font = New Font("Segoe UI", 9F, FontStyle.Bold)
    btnDeselectJobItems.Location = New Point(229, 383)
    btnDeselectJobItems.Name = "btnDeselectJobItems"
    btnDeselectJobItems.Size = New Size(51, 34)
    btnDeselectJobItems.TabIndex = 13
    btnDeselectJobItems.Text = "<<"
    btnDeselectJobItems.UseVisualStyleBackColor = True
    ' 
    ' btnClearAvailableJobs
    ' 
    btnClearAvailableJobs.Enabled = False
    btnClearAvailableJobs.Location = New Point(31, 460)
    btnClearAvailableJobs.Name = "btnClearAvailableJobs"
    btnClearAvailableJobs.Size = New Size(76, 34)
    btnClearAvailableJobs.TabIndex = 15
    btnClearAvailableJobs.Text = "Clear"
    btnClearAvailableJobs.UseVisualStyleBackColor = True
    ' 
    ' btnClearSelectedJobs
    ' 
    btnClearSelectedJobs.Enabled = False
    btnClearSelectedJobs.Location = New Point(306, 460)
    btnClearSelectedJobs.Name = "btnClearSelectedJobs"
    btnClearSelectedJobs.Size = New Size(76, 34)
    btnClearSelectedJobs.TabIndex = 16
    btnClearSelectedJobs.Text = "Clear"
    btnClearSelectedJobs.UseVisualStyleBackColor = True
    ' 
    ' Label5
    ' 
    Label5.AutoSize = True
    Label5.Location = New Point(511, 415)
    Label5.Name = "Label5"
    Label5.Size = New Size(122, 25)
    Label5.TabIndex = 18
    Label5.Text = "Project Name:"
    ' 
    ' txtProjectFilename
    ' 
    txtProjectFilename.Location = New Point(639, 412)
    txtProjectFilename.Name = "txtProjectFilename"
    txtProjectFilename.Size = New Size(173, 31)
    txtProjectFilename.TabIndex = 18
    ' 
    ' Label7
    ' 
    Label7.AutoSize = True
    Label7.Location = New Point(306, 213)
    Label7.Name = "Label7"
    Label7.Size = New Size(123, 25)
    Label7.TabIndex = 21
    Label7.Text = "Selected Jobs:"
    ' 
    ' btnGo
    ' 
    btnGo.Location = New Point(567, 460)
    btnGo.Name = "btnGo"
    btnGo.Size = New Size(112, 34)
    btnGo.TabIndex = 19
    btnGo.Text = "Go"
    btnGo.UseVisualStyleBackColor = True
    ' 
    ' btnClose
    ' 
    btnClose.Location = New Point(700, 460)
    btnClose.Name = "btnClose"
    btnClose.Size = New Size(112, 34)
    btnClose.TabIndex = 20
    btnClose.Text = "Close"
    btnClose.UseVisualStyleBackColor = True
    ' 
    ' Label8
    ' 
    Label8.AutoSize = True
    Label8.Location = New Point(510, 367)
    Label8.Name = "Label8"
    Label8.Size = New Size(111, 25)
    Label8.TabIndex = 24
    Label8.Text = "Puml Folder:"
    ' 
    ' txtPumlFolder
    ' 
    txtPumlFolder.Location = New Point(639, 364)
    txtPumlFolder.Name = "txtPumlFolder"
    txtPumlFolder.Size = New Size(111, 31)
    txtPumlFolder.TabIndex = 17
    txtPumlFolder.Text = "\PUML"
    ' 
    ' cbLibrary
    ' 
    cbLibrary.AutoSize = True
    cbLibrary.Location = New Point(99, 160)
    cbLibrary.Name = "cbLibrary"
    cbLibrary.Size = New Size(91, 29)
    cbLibrary.TabIndex = 4
    cbLibrary.Text = "Library"
    cbLibrary.UseVisualStyleBackColor = True
    ' 
    ' cbPDS
    ' 
    cbPDS.AutoSize = True
    cbPDS.Location = New Point(196, 160)
    cbPDS.Name = "cbPDS"
    cbPDS.Size = New Size(71, 29)
    cbPDS.TabIndex = 5
    cbPDS.Text = "PDS"
    cbPDS.UseVisualStyleBackColor = True
    ' 
    ' cbGDG
    ' 
    cbGDG.AutoSize = True
    cbGDG.Checked = True
    cbGDG.CheckState = CheckState.Checked
    cbGDG.Location = New Point(273, 160)
    cbGDG.Name = "cbGDG"
    cbGDG.Size = New Size(75, 29)
    cbGDG.TabIndex = 6
    cbGDG.Text = "GDG"
    cbGDG.UseVisualStyleBackColor = True
    ' 
    ' cbFile
    ' 
    cbFile.AutoSize = True
    cbFile.Checked = True
    cbFile.CheckState = CheckState.Checked
    cbFile.Location = New Point(354, 160)
    cbFile.Name = "cbFile"
    cbFile.Size = New Size(64, 29)
    cbFile.TabIndex = 7
    cbFile.Text = "File"
    cbFile.UseVisualStyleBackColor = True
    ' 
    ' Label9
    ' 
    Label9.AutoSize = True
    Label9.Location = New Point(20, 160)
    Label9.Name = "Label9"
    Label9.Size = New Size(73, 25)
    Label9.TabIndex = 30
    Label9.Text = "Include:"
    ' 
    ' lblJobsLoaded
    ' 
    lblJobsLoaded.AutoSize = True
    lblJobsLoaded.Location = New Point(34, 210)
    lblJobsLoaded.Name = "lblJobsLoaded"
    lblJobsLoaded.Size = New Size(141, 25)
    lblJobsLoaded.TabIndex = 31
    lblJobsLoaded.Text = "Jobs Loaded (0):"
    ' 
    ' cbSQL
    ' 
    cbSQL.AutoSize = True
    cbSQL.Checked = True
    cbSQL.CheckState = CheckState.Checked
    cbSQL.Location = New Point(424, 160)
    cbSQL.Name = "cbSQL"
    cbSQL.Size = New Size(70, 29)
    cbSQL.TabIndex = 32
    cbSQL.Text = "SQL"
    cbSQL.UseVisualStyleBackColor = True
    ' 
    ' Form1
    ' 
    AutoScaleDimensions = New SizeF(10F, 25F)
    AutoScaleMode = AutoScaleMode.Font
    ClientSize = New Size(836, 563)
    Controls.Add(cbSQL)
    Controls.Add(lblJobsLoaded)
    Controls.Add(Label9)
    Controls.Add(cbFile)
    Controls.Add(cbGDG)
    Controls.Add(cbPDS)
    Controls.Add(cbLibrary)
    Controls.Add(txtPumlFolder)
    Controls.Add(Label8)
    Controls.Add(btnClose)
    Controls.Add(btnGo)
    Controls.Add(Label7)
    Controls.Add(txtProjectFilename)
    Controls.Add(Label5)
    Controls.Add(btnClearSelectedJobs)
    Controls.Add(btnClearAvailableJobs)
    Controls.Add(btnDeselectJobItems)
    Controls.Add(btnDeselectJobItem)
    Controls.Add(btnSelectJobItems)
    Controls.Add(lblStatus)
    Controls.Add(btnSelectJobItem)
    Controls.Add(lbSelectedJobs)
    Controls.Add(lbAvailableJobs)
    Controls.Add(btnLoadJobs)
    Controls.Add(txtModelFilename)
    Controls.Add(Label4)
    Controls.Add(txtOutputFolder)
    Controls.Add(Label3)
    Controls.Add(Label2)
    Controls.Add(txtSandboxFolder)
    Controls.Add(Label1)
    Name = "Form1"
    Text = "Form1"
    ResumeLayout(False)
    PerformLayout()
  End Sub

  Friend WithEvents Label1 As Label
  Friend WithEvents txtSandboxFolder As TextBox
  Friend WithEvents Label2 As Label
  Friend WithEvents Label3 As Label
  Friend WithEvents txtOutputFolder As TextBox
  Friend WithEvents Label4 As Label
  Friend WithEvents txtModelFilename As TextBox
  Friend WithEvents btnLoadJobs As Button
  'Friend WithEvents lblJobsLoaded As Label
  Friend WithEvents lbAvailableJobs As ListBox
  Friend WithEvents lbSelectedJobs As ListBox
  Friend WithEvents btnSelectJobItem As Button
  Friend WithEvents lblStatus As Label
  Friend WithEvents btnSelectJobItems As Button
  Friend WithEvents btnDeselectJobItem As Button
  Friend WithEvents btnDeselectJobItems As Button
  Friend WithEvents btnClearAvailableJobs As Button
  Friend WithEvents btnClearSelectedJobs As Button
  Friend WithEvents Label5 As Label
  Friend WithEvents txtProjectFilename As TextBox
  'Friend WithEvents lblJobsLoaded As Label
  Friend WithEvents Label7 As Label
  Friend WithEvents btnGo As Button
  Friend WithEvents btnClose As Button
  Friend WithEvents Label8 As Label
  Friend WithEvents txtPumlFolder As TextBox
  Friend WithEvents cbLibrary As CheckBox
  Friend WithEvents cbPDS As CheckBox
  Friend WithEvents cbGDG As CheckBox
  Friend WithEvents cbFile As CheckBox
  Friend WithEvents Label9 As Label
  Friend WithEvents lblJobsLoaded As Label
  Friend WithEvents cbSQL As CheckBox

End Class
