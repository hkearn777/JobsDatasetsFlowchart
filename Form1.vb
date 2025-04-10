Imports System.IO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class Form1
  ' Create a Jobs with dataset flowchart with the PUML syntax.
  Dim ProgramVersion As String = "v0.0"
  'Change-history.
  ' 2025-03-18 v0.0 hk New code

  ' load the Excel References
  Dim objExcel As New Microsoft.Office.Interop.Excel.Application
  ' Model 
  Dim workbook As Microsoft.Office.Interop.Excel.Workbook
  Dim FilesWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim theWorksheet As Microsoft.Office.Interop.Excel.Worksheet


  Dim DefaultFormat = Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault
  Dim SetAsReadOnly = Microsoft.Office.Interop.Excel.XlFileAccess.xlReadOnly

  Dim Delimiter As String = "|"

  Dim dictJobStepsWithDatasets As New Dictionary(Of String, List(Of String))
  Dim Libraries As New List(Of String)
  Dim JobsUsingDatabase As New Dictionary(Of String, String)


  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Me.Text = "CreateJobsWithdatasetFlowchart " & ProgramVersion
    txtSandboxFolder.Text = Environment.GetEnvironmentVariable("ADDILite_Sandbox") &
      Environment.GetEnvironmentVariable("ADDILite_Application")

  End Sub
  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Me.Close()
  End Sub


  Private Sub btnLoadJobs_Click(sender As Object, e As EventArgs) Handles btnLoadJobs.Click
    ' load jobs to ListBox of available jobs from the model file
    ' verify user entries
    If Not VerifyUserEntries() Then
      Exit Sub
    End If

    Me.Cursor = Cursors.WaitCursor
    objExcel.Visible = False

    ' Open the Model spreadsheet
    Dim ModelFileName As String = txtSandboxFolder.Text & txtOutputFolder.Text & "\" & txtModelFilename.Text
    workbook = objExcel.Workbooks.Open(ModelFileName, True, SetAsReadOnly)

    ' load the libraries to the list of libraries
    Libraries = LoadLibraries()
    Libraries.Add("TEC1.LINKLIB")     'special hardcoded (boo!) library name

    ' load the jobs that use a database
    JobsUsingDatabase = LoadJobsUsingDatabase()

    ' load the jobs to the combo list
    dictJobStepsWithDatasets = LoadDictOfJobStepsWithDatasets()
    If dictJobStepsWithDatasets.Count = 0 Then
      MessageBox.Show("No Files found in the Model file")
    End If

    ' load the unique jobs to the Available Job list
    lblStatus.Text = "Loading Available Jobs"
    lbAvailableJobs.Items.Clear()
    For Each KeyItem In dictJobStepsWithDatasets.Keys
      Dim jobName As String = KeyItem.Split(Delimiter)(0)
      ' only add the job name once
      If Not lbAvailableJobs.Items.Contains(jobName) Then
        lbAvailableJobs.Items.Add(jobName)
      End If
    Next

    'close the Model spreadsheet
    workbook.Close()
    Me.Cursor = Cursors.Default

    ' default to first lbAvailableJobs entry
    If lbAvailableJobs.Items.Count > 0 Then
      lbAvailableJobs.SelectedIndex = 0
      lblStatus.Text = "Select Job Name(s) and click GO"
      lblJobsLoaded.Text = "Jobs Loaded (" & lbAvailableJobs.Items.Count & "):"
      btnSelectJobItem.Enabled = True
      btnSelectJobItems.Enabled = True
      btnDeselectJobItem.Enabled = True
      btnDeselectJobItems.Enabled = True
      lbAvailableJobs.Enabled = True
      lbSelectedJobs.Enabled = True
      btnClearSelectedJobs.Enabled = True
      btnClearAvailableJobs.Enabled = True
    End If
    lblStatus.Text = "Select Job Name(s) and click GO"
  End Sub
  Function LoadLibraries() As List(Of String)
    ' load the libraries from the model file
    Dim Libraries As New List(Of String)
    theWorksheet = workbook.Sheets.Item("Libraries")
    theWorksheet.Activate()
    Dim MaxRows As Long = theWorksheet.UsedRange.Rows(theWorksheet.UsedRange.Rows.Count).row
    Dim MaxCols As Long = theWorksheet.UsedRange.Columns.Count
    Dim StartRow As Long = 2
    For Row As Long = StartRow To MaxRows
      Dim LibraryName As String = GrabExcelField(Row, 1)
      If LibraryName.Trim.Length = 0 Then
        Continue For
      End If
      ' clean up the library name
      LibraryName = LibraryName.Trim.Replace("..", ".")
      ' check if libraryname already exists in the list
      If Libraries.Contains(LibraryName) Then
        Continue For
      End If
      Libraries.Add(LibraryName)
    Next
    Return Libraries
  End Function
  Function LoadDictOfJobStepsWithDatasets() As Dictionary(Of String, List(Of String))
    ' using the model load a list of job & steps names and datasets to the dictionary
    ' Assumption: the jobs are grouped together in the model file
    ' Assumption: the job sequence will be determined by the sequence they are loaded into the dictionary
    '             Maybe one day there will be a JOB Scheduler that will determine the sequence.

    lblStatus.Text = "Loading Jobs with Datasets"

    Dim dictofJobsandDatasets As New Dictionary(Of String, List(Of String))
    Dim listOfDatasets As New List(Of String)
    theWorksheet = workbook.Sheets.Item("Files")
    theWorksheet.Activate()
    Dim MaxRows As Long = theWorksheet.UsedRange.Rows(theWorksheet.UsedRange.Rows.Count).row
    Dim MaxCols As Long = theWorksheet.UsedRange.Columns.Count
    Dim jobNumber As Integer = 0
    Dim stepNumber As Integer = 0
    ' right justify and zero fill the integer job number into a string
    Dim jobNumberString As String = jobNumber.ToString.PadLeft(3, "0")
    ' right justify and zero fill the integer step number into a string
    Dim stepNumberString As String = ToString.PadLeft(3, "0")



    ' find the JOB
    Dim previousJob As String = GrabExcelField(2, 1)          'grab the first job name
    Dim previousStepName As String = GrabExcelField(2, 4)     'grab the first step name 
    Dim jobName As String = ""
    Dim stepName As String = ""

    Dim StartRow As Long = 2
    ' browse through the rows adding to the dictionary list
    For Row As Long = StartRow To MaxRows
      jobName = GrabExcelField(Row, 1)
      If jobName.Trim.Length = 0 Then
        Continue For
      End If
      If jobName = "CALLPGMS" Then
        Continue For
      End If
      If jobName = "ONLINE" Then
        Continue For
      End If
      stepName = GrabExcelField(Row, 4)
      If stepName.Trim.Length = 0 Then
        Continue For
      End If
      If stepName <> previousStepName Then
        stepNumber += 1
        stepNumberString = stepNumber.ToString.PadLeft(3, "0")
      End If
      If jobName <> previousJob Then
        If jobName <> previousJob Then
          jobNumber += 1
          jobNumberString = jobNumber.ToString.PadLeft(3, "0")
        End If
        listOfDatasets = AddSQLDatasettoDatasetList(previousJob, listOfDatasets)
        ' add the previous job to the dictionary
        dictofJobsandDatasets.Add(previousJob & Delimiter & jobNumberString & Delimiter &
                                  previousStepName & Delimiter & stepNumberString,
                                  listOfDatasets)
        ' empty the list of datasets 
        listOfDatasets = New List(Of String)
        previousJob = jobName
        previousStepName = stepName
        stepNumber = 0
      End If

      ' retrieve the dataset name
      Dim datasetName As String = GrabExcelField(Row, 10)
      ' filter out the 'not wanted' datasets (i.e. blank, or SYSOUT=, or Duplicate)
      If datasetName.Trim.Length = 0 Then
        Continue For
      End If
      datasetName = datasetName.Trim.Replace("..", ".")
      If datasetName.StartsWith("SYSOUT=") Then
        Continue For
      End If
      ' retrieve the DD field
      Dim ddField As String = GrabExcelField(Row, 7)
      If ddField.StartsWith("SORTWK") Then
        Continue For
      End If
      ' filter out the 'not wanted' DD names which will drop any unwanted datasets
      Select Case ddField
        Case "WORKSPACE", "CEEDUMP", "SYSOUT", "SYSUDUMP", "SYSPRINT", "SYSABEND", "SYSABOUT"
          Continue For
      End Select

      ' Retrieve start disp field
      Dim StartDispField As String = GrabExcelField(Row, 11)

      ' determing the dataset type by analyzing the dataset name and DD Name
      Dim datasetType As String = DetermineDatasetType(datasetName, ddField)

      ' drop if not selected dataset type
      If datasetType = "Library" And Not cbLibrary.Checked Then
        Continue For
      End If
      If datasetType = "PDS" And Not cbPDS.Checked Then
        Continue For
      End If
      If datasetType = "GDG" And Not cbGDG.Checked Then
        Continue For
      End If
      If datasetType = "File" And Not cbFile.Checked Then
        Continue For
      End If
      ' Note. sql datasets are dealt with in the LoadJobsUsingDatabase function


      Dim datasetInfo As String = datasetName & Delimiter &
                                  StartDispField & Delimiter &
                                  datasetType
      ' add the dataset to the list
      If Not listOfDatasets.Contains(datasetInfo) Then
        listOfDatasets.Add(datasetInfo)
      End If
    Next
    ' add the last job
    If listOfDatasets.Count > 0 Then
      jobNumber += 1
      jobNumberString = jobNumber.ToString.PadLeft(3, "0")
      stepNumber += 1
      stepNumberString = stepNumber.ToString.PadLeft(3, "0")
      dictofJobsandDatasets.Add(previousJob & Delimiter & jobNumberString & Delimiter &
                                previousStepName & stepNumberString,
                                listOfDatasets)
    End If
    Return dictofJobsandDatasets
  End Function
  Function AddSQLDatasettoDatasetList(ByRef theJob As String,
                                      ByRef listofdatasets As List(Of String)) As List(Of String)
    ' this function will add the SQL dataset to the list of datasets if requested to do so and 
    ' if the job has any SQL datasets. If the SQL dataset have many open types, then
    ' the SQL dataset will be added to the list of datasets for each open type.
    ' Also need to convert the SQL open types to either INPUT or OUTPUT values.
    '   SELECT = INPUT, UPDATE = OUTPUT, INSERT = OUTPUT, DELETE = OUTPUT, CURSOR = INPUT
    If cbSQL.Checked Then
      If JobsUsingDatabase.ContainsKey(theJob) Then
        Dim myValue As String = JobsUsingDatabase.Item(theJob)
        Dim myValues() As String = myValue.Split(Delimiter)
        For currentDataset As Integer = 0 To myValues.Count - 1 Step 3
          Dim programName As String = myValues(currentDataset)
          Dim tableName As String = myValues(currentDataset + 1)
          Dim openType As String = myValues(currentDataset + 2)
          Dim openTypes() As String = openType.Split(" "c)
          ' process each type of open type (i.e. SELECT, INSERT, UPDATE, DELETE)
          For Each type As String In openTypes
            If type.Trim.Length = 0 Then
              Continue For
            End If
            ' determine the start disp field
            Dim StartDispField As String = "INPUT"
            If type = "UPDATE" Or type = "INSERT" Or type = "DELETE" Then
              StartDispField = "OUTPUT"
            End If
            ' check if the dataset name already exists in the list
            Dim mydatasetInfo As String = tableName & Delimiter & StartDispField & Delimiter & "SQL"
            If Not listofdatasets.Contains(mydatasetInfo) Then
              listofdatasets.Add(mydatasetInfo)
            End If
          Next
        Next
      End If
    End If

    Return listofdatasets
  End Function

  Function DetermineDatasetType(ByRef datasetName As String, ByRef ddField As String) As String
    ' determine the dataset type by analyzing the dataset name and/or DD field
    ' types of datasets are:  PDS, GDG, File, Library (object,load,steplib,joblib,etc.), SQL
    ' if the dataset name is not recognized, then it is classified as "Unknown"
    ' The check for SQL is done in the LoadJobsUsingDatabase function
    If ddField = "STEPLIB" Or ddField = "JOBLIB" Then
      Return "Library"
    End If
    ' check if datasetname is in the library list
    If Libraries.Contains(datasetName) Then
      Return "Library"
    End If
    If datasetName.Contains("()") Then
      Return "GDG"
    End If

    ' check for PDS; PDS's have a (???) between the parenthesis
    Dim startOpenParen As Integer = datasetName.IndexOf("("c)
    Dim endOpenParen As Integer = datasetName.IndexOf(")"c)
    If startOpenParen <= 0 Or endOpenParen <= 0 Then
    Else
      If endOpenParen > startOpenParen + 1 Then
        Return "PDS"
      End If
    End If

    ' check for File
    If datasetName.Contains("."c) Then
      Return "File"
    End If
    Return "Unknown"
  End Function

  Function GrabExcelField(ByRef theRow As Integer, ByRef theColumn As Integer) As String
    If theRow = 0 Then
      Return ""
    End If
    If theColumn = 0 Then
      Return ""
    End If
    Dim theValue As String = theWorksheet.Cells(theRow, theColumn).value2
    If theValue Is Nothing Then
      Return ""
    End If
    If theValue.Length = 0 Then
      Return ""
    End If
    Return theValue
  End Function

  Function VerifyUserEntries() As Boolean

    ' verify file names
    If Not Directory.Exists(txtSandboxFolder.Text) Then
      MessageBox.Show("Sandbox folder not found")
      Return False
    End If

    ' verify output folder name has a value
    If txtOutputFolder.Text.Trim.Length = 0 Then
      MessageBox.Show("Output folder is required")
      Return False
    End If

    ' verify output folder
    Dim outputFolder As String = txtSandboxFolder.Text & txtOutputFolder.Text
    If Not Directory.Exists(outputFolder) Then
      MessageBox.Show("Output folder not found:" & vbCrLf & outputFolder)
      Return False
    End If

    ' verify model file exists Note. the model resides in the output folder
    Dim ModelFileName As String = outputFolder & "\" & txtModelFilename.Text
    If Not File.Exists(ModelFileName) Then
      MessageBox.Show("Model file not found" & vbCrLf & ModelFileName)
      Return False
    End If
    Return True
  End Function

  Private Sub btnSelectJobItem_Click(sender As Object, e As EventArgs) Handles btnSelectJobItem.Click
    ' ensure a job is selected
    If lbAvailableJobs.SelectedIndex = -1 Then
      MessageBox.Show("Select a Job")
      Exit Sub
    End If
    ' move the selected job to the selected jobs list
    lbSelectedJobs.Items.Add(lbAvailableJobs.SelectedItem)
  End Sub

  Private Sub btnSelectJobItems_Click(sender As Object, e As EventArgs) Handles btnSelectJobItems.Click
    ' ensure one or more jobs are selected
    If lbAvailableJobs.SelectedIndex = -1 Then
      MessageBox.Show("Select a Job")
      Exit Sub
    End If
    ' move the selected jobs to the selected jobs list
    For Each item In lbAvailableJobs.SelectedItems
      lbSelectedJobs.Items.Add(item)
    Next

  End Sub

  Private Sub btnDeselectJobItem_Click(sender As Object, e As EventArgs) Handles btnDeselectJobItem.Click
    ' ensure a job is selected in the selected jobs list
    If lbSelectedJobs.SelectedIndex = -1 Then
      MessageBox.Show("Select a Job")
      Exit Sub
    End If
    ' remove the selected job from the selected jobs list
    lbSelectedJobs.Items.RemoveAt(lbSelectedJobs.SelectedIndex)

  End Sub

  Private Sub btnDeselectJobItems_Click(sender As Object, e As EventArgs) Handles btnDeselectJobItems.Click
    ' ensure one or more jobs are selected
    If lbSelectedJobs.SelectedIndex = -1 Then
      MessageBox.Show("Select a Job")
      Exit Sub
    End If
    ' save current selected items
    Dim savedItems As New List(Of String)
    For Each item In lbSelectedJobs.Items
      savedItems.Add(item)
    Next
    ' grab the selected items and store in itemsToRemove list
    Dim itemsToRemove As New List(Of String)
    For Each item In lbSelectedJobs.SelectedItems
      itemsToRemove.Add(item)
    Next
    ' remove items from the list
    For Each item In itemsToRemove
      savedItems.Remove(item)
    Next
    ' now reload the lblSelectedJobs list from the remaining items in the savedItems list
    lbSelectedJobs.Items.Clear()
    For Each item In savedItems
      lbSelectedJobs.Items.Add(item)
    Next
  End Sub

  Private Sub btnClearSelectedJobs_Click(sender As Object, e As EventArgs) Handles btnClearSelectedJobs.Click
    ' clear the selected jobs list
    lbSelectedJobs.Items.Clear()
  End Sub

  Private Sub btnClearAvailableJobs_Click(sender As Object, e As EventArgs) Handles btnClearAvailableJobs.Click
    ' clear the available jobs list
    lbAvailableJobs.Items.Clear()
    btnSelectJobItem.Enabled = False
    btnSelectJobItems.Enabled = False
    btnDeselectJobItem.Enabled = False
    btnDeselectJobItems.Enabled = False
    lbAvailableJobs.Enabled = False
    lbSelectedJobs.Enabled = False
    btnClearSelectedJobs.Enabled = False
    btnClearAvailableJobs.Enabled = False
    ' clear the selected jobs list
    lbSelectedJobs.Items.Clear()
  End Sub


  Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
    ' are there any jobs selected
    If lbSelectedJobs.Items.Count = 0 Then
      MessageBox.Show("Select one or more Jobs")
      Exit Sub
    End If
    ' verify project file name is entered
    If txtProjectFilename.Text.Trim.Length = 0 Then
      MessageBox.Show("Enter a Project File Name")
      Exit Sub
    End If
    ' verify puml folder name is entered
    If txtPumlFolder.Text.Trim.Length = 0 Then
      MessageBox.Show("Enter a PUML Folder Name")
      Exit Sub
    End If
    ' verify puml folder exists
    Dim pumlFolder As String = txtSandboxFolder.Text & txtPumlFolder.Text
    If Not Directory.Exists(pumlFolder) Then
      MessageBox.Show("PUML folder not found:" & vbCrLf & pumlFolder)
      Exit Sub
    End If

    Me.Cursor = Cursors.WaitCursor

    Call CreateFlowcharts()

    Me.Cursor = Cursors.Default
    MessageBox.Show("Flowchart Puml(s) created")
  End Sub
  Sub CreateFlowcharts()
    ' this routine will create the flowchart for the job(s) selected

    Dim JobsAndUniqueDatasets As List(Of String) = LoadListOfJobsWithDatasets()

    ' define the project file name
    Dim ProjectFileName As String = txtSandboxFolder.Text &
                                    txtPumlFolder.Text &
                                    "\" & txtProjectFilename.Text &
                                    ".puml"

    ' open the PUML project file text file
    Dim swPuml As StreamWriter = New StreamWriter(ProjectFileName, False)
    swPuml.WriteLine("@startuml " & txtProjectFilename.Text & ".puml")
    swPuml.WriteLine("skinparam shadowing false")
    swPuml.WriteLine("header " & Me.Text & "(c), by IBM")
    swPuml.WriteLine("footer " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    ' append the check box fields to the title
    Dim checkBoxFields As String = ""
    If cbLibrary.Checked Then
      checkBoxFields = checkBoxFields & "Library "
    End If
    If cbPDS.Checked Then
      checkBoxFields = checkBoxFields & "PDS "
    End If
    If cbGDG.Checked Then
      checkBoxFields = checkBoxFields & "GDG "
    End If
    If cbFile.Checked Then
      checkBoxFields = checkBoxFields & "File "
    End If
    If cbSQL.Checked Then
      checkBoxFields = checkBoxFields & "SQL "
    End If
    swPuml.WriteLine("title Jobs with Datasets\nDatasets selected: " & checkBoxFields)
    swPuml.WriteLine("")
    swPuml.WriteLine("left to right direction")
    swPuml.WriteLine("")

    Dim jobName As String = ""
    For Each jobName In lbSelectedJobs.Items
      createflowchart(jobName, swPuml, JobsAndUniqueDatasets)
    Next

    ' close the project file
    swPuml.WriteLine("@enduml")
    swPuml.Close()

  End Sub
  Sub createflowchart(ByRef jobName As String,
                      ByRef swPuml As StreamWriter,
                      ByRef JobsAndUniqueDatasets As List(Of String))
    ' this routine will create the flowchart puml details for a job selected

    ' puml syntax colors and lines
    '  #line:red;line.dotted;text:red when the dataset/object is an input and output
    ' dsn -[#blue]down-> job... when the dataset/object is an input
    ' job -[#green]down-> dsn... when the dataset/object is an output
    ' left to right direction
    ' job <-[#red]> app... when the dataset/object is an input and output (left side of object)
    ' job <-[#red]left-> app... when the dataset/object is an input and output (right side of object)
    ' top to bottom direction (back to default)


    ' all the datasets for the job
    Dim theJobName As String
    Dim seqNumber As String
    Dim datasetName As String
    Dim StartDispField As String
    Dim datasetType As String

    ' build the objects for the job (jobname and files)
    swPuml.WriteLine("rectangle " & jobName)
    For Each datasetInfo In JobsAndUniqueDatasets
      theJobName = datasetInfo.Split(Delimiter)(0)
      If theJobName <> jobName Then
        Continue For
      End If
      seqNumber = datasetInfo.Split(Delimiter)(1)
      datasetName = datasetInfo.Split(Delimiter)(2)
      StartDispField = datasetInfo.Split(Delimiter)(3)
      datasetType = datasetInfo.Split(Delimiter)(4)
      Dim color As String = ""
      Dim objType As String = "file"
      Select Case datasetType
        Case "PDS"
          objType = "collections"
          color = "#LightGreen"
        Case "GDG"
          objType = "collections"
          color = "#LightGreen"
        Case "Library"
          objType = "folder"
          color = "#LightBlue"
        Case "SQL"
          objType = "database"
          color = "#LightYellow"
      End Select
      swPuml.WriteLine(objType & " " & Chr(34) & datasetName & Chr(34) & " " & Color)
    Next
    swPuml.WriteLine("")

    ' map all the inputs to the job
    For Each datasetInfo In JobsAndUniqueDatasets
      theJobName = datasetInfo.Split(Delimiter)(0)
      If theJobName <> jobName Then
        Continue For
      End If
      seqNumber = datasetInfo.Split(Delimiter)(1)
      datasetName = datasetInfo.Split(Delimiter)(2)
      StartDispField = datasetInfo.Split(Delimiter)(3)
      datasetType = datasetInfo.Split(Delimiter)(4)
      If StartDispField = "INPUT" Then
        swPuml.WriteLine(Chr(34) & datasetName & Chr(34) & " -[#blue]-> " & jobName)
      End If
    Next
    swPuml.WriteLine("")

    ' map all the datasets that are both input and output to the job
    For Each datasetInfo In JobsAndUniqueDatasets
      theJobName = datasetInfo.Split(Delimiter)(0)
      If theJobName <> jobName Then
        Continue For
      End If
      seqNumber = datasetInfo.Split(Delimiter)(1)
      datasetName = datasetInfo.Split(Delimiter)(2)
      StartDispField = datasetInfo.Split(Delimiter)(3)
      datasetType = datasetInfo.Split(Delimiter)(4)
      If StartDispField = "BOTH" Then
        swPuml.WriteLine(jobName & " <-[#red]> " & Chr(34) & datasetName & Chr(34))
      End If
    Next


    ' map all the outputs to the job
    For Each datasetInfo In JobsAndUniqueDatasets
      theJobName = datasetInfo.Split(Delimiter)(0)
      If theJobName <> jobName Then
        Continue For
      End If
      seqNumber = datasetInfo.Split(Delimiter)(1)
      datasetName = datasetInfo.Split(Delimiter)(2)
      StartDispField = datasetInfo.Split(Delimiter)(3)
      datasetType = datasetInfo.Split(Delimiter)(4)
      If StartDispField = "OUTPUT" Then
        swPuml.WriteLine(jobName & " -[#green]-> " & Chr(34) & datasetName & Chr(34))
      End If
    Next
    swPuml.WriteLine("")


  End Sub
  Function LoadListOfJobsWithDatasets() As List(Of String)
    ' this function will analyze the dictofJobStepWithDatasets and return a dictionary of datasets
    ' by job / dataset. Using the start disp field to determine
    ' if the datasets are input or output or both to each job.


    Dim dictOfJobsWithDatasets As New Dictionary(Of String, String) ' key=jobname|sequence, value=datasetname|startdisp

    ' need to create a list of dataset names for each job and then sort by jobname and dataset name
    Dim listOfJobsAndDatasets As New List(Of String)
    For Each JobKey In dictJobStepsWithDatasets.Keys
      Dim myJobName As String = JobKey.Split(Delimiter)(0)
      ' see if jobname is in the selected jobs list
      If Not lbSelectedJobs.Items.Contains(myJobName) Then
        Continue For
      End If
      Dim myJobSequence As String = JobKey.Split(Delimiter)(1)
      Dim myStepName As String = JobKey.Split(Delimiter)(2)
      Dim myStepSequence As String = JobKey.Split(Delimiter)(3)
      Dim listOfDatasets As New List(Of String)
      listOfDatasets = dictJobStepsWithDatasets.Item(JobKey)
      For Each datasetInfo As String In listOfDatasets
        Dim myDatasetName As String = datasetInfo.Split(Delimiter)(0)
        Dim myStartDispField As String = datasetInfo.Split(Delimiter)(1)
        Dim myDatasetType As String = datasetInfo.Split(Delimiter)(2)
        listOfJobsAndDatasets.Add(myJobName & Delimiter & myJobSequence & Delimiter &
                                  myDatasetName & Delimiter &
                                  myStartDispField & Delimiter &
                                  myDatasetType)
      Next
    Next
    listOfJobsAndDatasets.Sort()

    ' now I need analyze the startdisp fields; if INPUT and OUTPUT then make it BOTH

    ' count the number of unique datasets for a jobname
    Dim dictOfDatasets As New Dictionary(Of String, Integer)
    Dim listOfJobsAndUniqueDatasets As New List(Of String)
    Dim jobName(1) As String
    Dim jobSequence(1) As String
    Dim datasetName(1) As String
    Dim startdisp(1) As String
    Dim datasetType(1) As String
    For item As Integer = 0 To listOfJobsAndDatasets.Count - 1 ' process in pairs...
      If item = listOfJobsAndDatasets.Count - 1 Then
        Exit For
      End If
      ' current index
      jobName(0) = listOfJobsAndDatasets(item).Split(Delimiter)(0)
      jobSequence(0) = listOfJobsAndDatasets(item).Split(Delimiter)(1)
      datasetName(0) = listOfJobsAndDatasets(item).Split(Delimiter)(2)
      startdisp(0) = listOfJobsAndDatasets(item).Split(Delimiter)(3)
      datasetType(0) = listOfJobsAndDatasets(item).Split(Delimiter)(4)
      ' next index
      jobName(1) = listOfJobsAndDatasets(item + 1).Split(Delimiter)(0)
      jobSequence(1) = listOfJobsAndDatasets(item + 1).Split(Delimiter)(1)
      datasetName(1) = listOfJobsAndDatasets(item + 1).Split(Delimiter)(2)
      startdisp(1) = listOfJobsAndDatasets(item + 1).Split(Delimiter)(3)
      datasetType(1) = listOfJobsAndDatasets(item + 1).Split(Delimiter)(4)

      If jobName(0) = jobName(1) And datasetName(0) = datasetName(1) Then
        If startdisp(0) <> startdisp(1) Then
          startdisp(0) = "BOTH"
        End If
        item += 1
      End If
      listOfJobsAndUniqueDatasets.Add(jobName(0) & Delimiter &
                                      jobSequence(0) & Delimiter &
                                      datasetName(0) & Delimiter &
                                      startdisp(0) & Delimiter &
                                      datasetType(0))
    Next

    Return listOfJobsAndUniqueDatasets
  End Function
  Function LoadJobsUsingDatabase() As Dictionary(Of String, String)

    lblStatus.Text = "Loading Jobs that uses Databases"

    ' load the programs that use a database
    Dim myJobsUsingDatabase As New Dictionary(Of String, String)

    theWorksheet = workbook.Sheets.Item("Records")
    theWorksheet.Activate()

    ' Apply filter to the Type column (E) for a specific value of SQL
    'theWorksheet.Range("E1").AutoFilter(Field:=5, Criteria1:="SQL")
    'Dim usedRange As Microsoft.Office.Interop.Excel.Range = theWorksheet.UsedRange

    ' Apply advanced filter to column B to get unique values
    'Dim uniqueRange As Microsoft.Office.Interop.Excel.Range = theWorksheet.Range("B1", "B" & usedRange.Rows.Count)
    'uniqueRange.AdvancedFilter(Action:=Microsoft.Office.Interop.Excel.XlFilterAction.xlFilterInPlace, Unique:=True)

    Dim MaxRows As Long = theWorksheet.UsedRange.Rows(theWorksheet.UsedRange.Rows.Count).row
    Dim MaxCols As Long = theWorksheet.UsedRange.Columns.Count
    Dim StartRow As Long = 2

    ' Browse through the rows adding to the programsUsingDatabase list    
    For Row As Long = StartRow To MaxRows
      Dim theType As String = GrabExcelField(Row, 5)
      If theType.Trim.Length = 0 Then
        Continue For
      End If
      If theType <> "SQL" Then
        Continue For
      End If
      Dim theJobName As String = GrabExcelField(Row, 1)
      If theJobName.Trim.Length = 0 Then
        Continue For
      End If
      Dim theProgram As String = GrabExcelField(Row, 2)
      Dim theTable As String = GrabExcelField(Row, 3)
      Dim theOpen As String = GrabExcelField(Row, 11)
      If theOpen.Trim.Length = 0 Then
        theOpen = "SELECT"
      End If
      Dim theValue As String = theProgram & Delimiter & theTable & Delimiter & theOpen
      ' Check if the key already exists in the dictionary
      If Not myJobsUsingDatabase.ContainsKey(theJobName) Then
        myJobsUsingDatabase.Add(theJobName, theValue)
      Else
        ' If the key exists, append the new value to the existing value
        myJobsUsingDatabase(theJobName) &= Delimiter & theValue
      End If

    Next

    'For Each row As Microsoft.Office.Interop.Excel.Range In usedRange.Rows
    '  If row.EntireRow.Hidden = False Then
    '    ' Read the value from the first column (A)
    '    Dim theKey As String = row.Cells(1, 1).Value2.ToString()
    '    If theKey = "Source" Then
    '      Continue For
    '    End If
    '    Dim theProgram As String =
    '      row.Cells(1, 2).Value2.ToString()
    '    Dim theTable As String = row.Cells(1, 3).Value2.ToString()
    '    Dim theOpen As String = If(row.Cells(1, 11).Value2, "SELECT")
    '    Dim theValue As String = theProgram & Delimiter & theTable & Delimiter & theOpen
    '    'Dim theValue As String = theTable & Delimiter & theOpen
    '    ' Check if the key already exists in the dictionary
    '    If Not myJobsUsingDatabase.ContainsKey(theKey) Then
    '      myJobsUsingDatabase.Add(theKey, theValue)
    '    End If
    '  End If
    'Next

    ' Remove the filters
    'theWorksheet.AutoFilterMode = False

    Return myJobsUsingDatabase
  End Function
End Class
