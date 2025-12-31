Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Shapes
Imports Newtonsoft.Json

Class MainWindow

  ' Store the loaded diagram data
  Private currentDiagramData As DiagramData = Nothing

  ' ========== Zoom and Pan fields ==========
  Private zoomLevel As Double = 1.0
  Private Const MinZoom As Double = 0.1
  Private Const MaxZoom As Double = 5.0
  Private Const ZoomStep As Double = 0.2

  ' Pan dragging fields
  Private isPanning As Boolean = False
  Private panStartPoint As Point

  ' ========== NEW: Search fields ==========
  Private searchResults As New List(Of SearchResultItem)
  Private currentSearchIndex As Integer = -1
  Private originalBrushes As New Dictionary(Of Rectangle, Brush)

  ' ========== Layout constants (moved to class level) ==========
  Private Const JOB_WIDTH As Double = 91
  Private Const JOB_HEIGHT As Double = 38
  Private Const DATASET_WIDTH As Double = 120
  Private Const DATASET_HEIGHT As Double = 51
  Private Const SPACING_Y As Double = 15          ' Spacing between datasets
  Private Const SPACING_Y_DATABASE As Double = 25 ' Extra spacing for database shapes
  Private Const START_X As Double = 30
  Private Const START_Y As Double = 30
  Private Const MIN_JOB_SPACING As Double = 40    ' Minimum space between jobs
  Private Const COLUMN_GAP As Double = 220        ' Gap between columns (increased from 180 to 220)
  Private Const JOB_TO_BOTH_GAP As Double = 80    ' Gap between job and its BOTH datasets
  Private Const BOTH_HORIZONTAL_OFFSET As Double = 40  ' Shift BOTH datasets right for connection space

  ' Helper class to store search results
  Private Class SearchResultItem
    Public Property Rectangle As Rectangle
    Public Property DatasetInfo As DatasetInfo
    Public Property JobName As String
    Public Property Position As Point
  End Class

  ' ========== Zoom Control Event Handlers ==========

  Private Sub btnZoomIn_Click(sender As Object, e As RoutedEventArgs)
    ZoomBy(ZoomStep)
  End Sub

  Private Sub btnZoomOut_Click(sender As Object, e As RoutedEventArgs)
    ZoomBy(-ZoomStep)
  End Sub

  Private Sub btnResetView_Click(sender As Object, e As RoutedEventArgs)
    ' Reset zoom and pan to default
    zoomLevel = 1.0
    scaleTransform.ScaleX = zoomLevel
    scaleTransform.ScaleY = zoomLevel
    translateTransform.X = 0
    translateTransform.Y = 0
    UpdateZoomLabel()
    txtStatusBar.Text = "View reset to 100%"
  End Sub

  Private Sub ZoomBy(delta As Double)
    Dim newZoom As Double = Math.Max(MinZoom, Math.Min(MaxZoom, zoomLevel + delta))

    If newZoom <> zoomLevel Then
      ' Get the center point of the viewport for smooth zooming
      Dim centerX As Double = scrollViewer.ViewportWidth / 2
      Dim centerY As Double = scrollViewer.ViewportHeight / 2

      ' Calculate the zoom factor
      Dim factor As Double = newZoom / zoomLevel

      ' Adjust translate to keep center point stable
      translateTransform.X = (translateTransform.X - centerX) * factor + centerX
      translateTransform.Y = (translateTransform.Y - centerY) * factor + centerY

      ' Update zoom level
      zoomLevel = newZoom
      scaleTransform.ScaleX = zoomLevel
      scaleTransform.ScaleY = zoomLevel

      UpdateZoomLabel()
      txtStatusBar.Text = $"Zoom: {Math.Round(zoomLevel * 100)}%"
    End If
  End Sub

  Private Sub CanvasContainer_MouseWheel(sender As Object, e As MouseWheelEventArgs)
    ' Ctrl+MouseWheel for zoom
    If Keyboard.Modifiers = ModifierKeys.Control Then
      Dim delta As Double = If(e.Delta > 0, ZoomStep, -ZoomStep)

      ' Zoom towards mouse cursor position
      Dim mousePos As Point = e.GetPosition(canvasContainer)
      ZoomToPoint(delta, mousePos)

      e.Handled = True
    End If
  End Sub

  Private Sub ZoomToPoint(delta As Double, point As Point)
    Dim newZoom As Double = Math.Max(MinZoom, Math.Min(MaxZoom, zoomLevel + delta))

    If newZoom <> zoomLevel Then
      Dim factor As Double = newZoom / zoomLevel

      ' Zoom towards the specified point
      translateTransform.X = (translateTransform.X - point.X) * factor + point.X
      translateTransform.Y = (translateTransform.Y - point.Y) * factor + point.Y

      zoomLevel = newZoom
      scaleTransform.ScaleX = zoomLevel
      scaleTransform.ScaleY = zoomLevel

      UpdateZoomLabel()
      txtStatusBar.Text = $"Zoom: {Math.Round(zoomLevel * 100)}%"
    End If
  End Sub

  Private Sub UpdateZoomLabel()
    txtZoomLevel.Text = $"{Math.Round(zoomLevel * 100)}%"
  End Sub

  ' ========== Pan Control Event Handlers ==========

  Private Sub CanvasContainer_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
    ' Start panning only if clicking on empty space (not on a dataset rectangle)
    If TypeOf e.OriginalSource Is Grid OrElse TypeOf e.OriginalSource Is Canvas Then
      isPanning = True
      panStartPoint = e.GetPosition(scrollViewer)
      canvasContainer.Cursor = Cursors.Hand
      canvasContainer.CaptureMouse()
      e.Handled = True
    End If
  End Sub

  Private Sub CanvasContainer_MouseMove(sender As Object, e As MouseEventArgs)
    If isPanning Then
      Dim currentPoint As Point = e.GetPosition(scrollViewer)
      Dim deltaX As Double = currentPoint.X - panStartPoint.X
      Dim deltaY As Double = currentPoint.Y - panStartPoint.Y

      translateTransform.X += deltaX
      translateTransform.Y += deltaY

      panStartPoint = currentPoint
      e.Handled = True
    End If
  End Sub

  Private Sub CanvasContainer_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)
    If isPanning Then
      isPanning = False
      canvasContainer.Cursor = Cursors.Arrow
      canvasContainer.ReleaseMouseCapture()
      e.Handled = True
    End If
  End Sub

  ' ========== NEW: Search Functionality ==========

  Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs)
    PerformSearch()
  End Sub

  Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs)
    ' Search when Enter key is pressed
    If e.Key = Key.Enter Then
      PerformSearch()
      e.Handled = True
    End If
  End Sub

  Private Sub txtSearch_TextChanged(sender As Object, e As TextChangedEventArgs)
    ' Clear search if text box is empty
    If String.IsNullOrWhiteSpace(txtSearch.Text) Then
      ClearSearch()
    End If
  End Sub

  Private Sub btnPrevious_Click(sender As Object, e As RoutedEventArgs)
    NavigateSearchResults(-1)
  End Sub

  Private Sub btnNext_Click(sender As Object, e As RoutedEventArgs)
    NavigateSearchResults(1)
  End Sub

  Private Sub btnClearSearch_Click(sender As Object, e As RoutedEventArgs)
    ClearSearch()
    txtSearch.Clear()
  End Sub

  Private Sub PerformSearch()
    ' Clear previous search results
    ClearSearch()

    Dim searchTerm As String = txtSearch.Text.Trim()

    If String.IsNullOrWhiteSpace(searchTerm) Then
      txtStatusBar.Text = "Enter a search term"
      Return
    End If

    ' Search through all rectangles on the canvas
    For Each child In diagramCanvas.Children
      If TypeOf child Is Rectangle Then
        Dim rect As Rectangle = CType(child, Rectangle)

        ' Check if this rectangle has a Tag (datasets have tags, jobs don't in current implementation)
        If rect.Tag IsNot Nothing AndAlso TypeOf rect.Tag Is DatasetInfo Then
          Dim dataset As DatasetInfo = CType(rect.Tag, DatasetInfo)

          ' Check if dataset name contains search term (case-insensitive)
          If dataset.name.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0 Then
            ' Store original brush before highlighting
            If Not originalBrushes.ContainsKey(rect) Then
              originalBrushes.Add(rect, rect.Fill)
            End If

            ' Get position for navigation
            Dim rectPos As New Point(Canvas.GetLeft(rect), Canvas.GetTop(rect))

            ' Add to search results
            searchResults.Add(New SearchResultItem With {
              .Rectangle = rect,
              .DatasetInfo = dataset,
              .Position = rectPos
            })
          End If
        End If
      End If
    Next

    ' Also search job names (check TextBlocks)
    For Each child In diagramCanvas.Children
      If TypeOf child Is TextBlock Then
        Dim textBlock As TextBlock = CType(child, TextBlock)

        ' Job labels have FontWeight.Bold and larger font
        If textBlock.FontWeight = FontWeights.Bold AndAlso textBlock.FontSize = 14 Then
          If textBlock.Text.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0 Then
            ' Find the associated rectangle (jobs don't have Tags)
            Dim textPos = New Point(Canvas.GetLeft(textBlock), Canvas.GetTop(textBlock))

            ' Search for rectangle at same position (job rectangle is behind the text)
            For Each child2 In diagramCanvas.Children
              If TypeOf child2 Is Rectangle Then
                Dim rect As Rectangle = CType(child2, Rectangle)
                If rect.Tag Is Nothing Then ' Job rectangles don't have tags
                  Dim rectPos As Point = New Point(Canvas.GetLeft(rect), Canvas.GetTop(rect))
                  ' Check if positions match (allowing small tolerance)
                  If Math.Abs(rectPos.X - textPos.X) < 10 AndAlso Math.Abs(rectPos.Y - (textPos.Y - rect.Height / 2 + 10)) < 10 Then
                    ' Store original brush
                    If Not originalBrushes.ContainsKey(rect) Then
                      originalBrushes.Add(rect, rect.Fill)
                    End If

                    searchResults.Add(New SearchResultItem With {
                      .Rectangle = rect,
                      .JobName = textBlock.Text,
                      .Position = rectPos
                    })
                    Exit For
                  End If
                End If
              End If
            Next
          End If
        End If
      End If
    Next

    ' Update UI based on results
    If searchResults.Count > 0 Then
      ' Highlight all matches
      For Each result In searchResults
        result.Rectangle.Fill = Brushes.Yellow
        result.Rectangle.StrokeThickness = 3
      Next

      ' Navigate to first result
      currentSearchIndex = 0
      HighlightCurrentResult()
      NavigateToCurrentResult()

      txtSearchResults.Text = $"1 of {searchResults.Count}"
      txtStatusBar.Text = $"Found {searchResults.Count} match(es) for '{searchTerm}'"
    Else
      txtSearchResults.Text = "0 results"
      txtStatusBar.Text = $"No matches found for '{searchTerm}'"
    End If
  End Sub

  Private Sub NavigateSearchResults(direction As Integer)
    If searchResults.Count = 0 Then
      Return
    End If

    ' Move index
    currentSearchIndex += direction

    ' Wrap around
    If currentSearchIndex < 0 Then
      currentSearchIndex = searchResults.Count - 1
    ElseIf currentSearchIndex >= searchResults.Count Then
      currentSearchIndex = 0
    End If

    HighlightCurrentResult()
    NavigateToCurrentResult()

    txtSearchResults.Text = $"{currentSearchIndex + 1} of {searchResults.Count}"
  End Sub

  Private Sub HighlightCurrentResult()
    ' Reset all to yellow
    For Each result In searchResults
      result.Rectangle.Fill = Brushes.Yellow
      result.Rectangle.StrokeThickness = 3
    Next

    ' Highlight current result in orange
    If currentSearchIndex >= 0 AndAlso currentSearchIndex < searchResults.Count Then
      Dim currentResult = searchResults(currentSearchIndex)
      currentResult.Rectangle.Fill = Brushes.Orange
      currentResult.Rectangle.StrokeThickness = 4
    End If
  End Sub

  Private Sub NavigateToCurrentResult()
    If currentSearchIndex < 0 OrElse currentSearchIndex >= searchResults.Count Then
      Return
    End If

    Dim result = searchResults(currentSearchIndex)

    ' Calculate center point of the result rectangle
    Dim centerX As Double = result.Position.X + result.Rectangle.Width / 2
    Dim centerY As Double = result.Position.Y + result.Rectangle.Height / 2

    ' Calculate viewport center
    Dim viewportCenterX As Double = scrollViewer.ViewportWidth / 2
    Dim viewportCenterY As Double = scrollViewer.ViewportHeight / 2

    ' Pan to center the result (accounting for current zoom and transform)
    translateTransform.X = viewportCenterX - (centerX * zoomLevel)
    translateTransform.Y = viewportCenterY - (centerY * zoomLevel)

    ' Update status with item info
    Dim itemName As String = If(result.DatasetInfo IsNot Nothing, result.DatasetInfo.name, result.JobName)
    txtStatusBar.Text = $"Showing result {currentSearchIndex + 1} of {searchResults.Count}: {itemName}"
  End Sub

  Private Sub ClearSearch()
    ' Restore original colors
    For Each result In searchResults
      If originalBrushes.ContainsKey(result.Rectangle) Then
        result.Rectangle.Fill = originalBrushes(result.Rectangle)
        result.Rectangle.StrokeThickness = If(result.DatasetInfo IsNot Nothing, 1.5, 2)
      End If
    Next

    ' Clear collections
    searchResults.Clear()
    originalBrushes.Clear()
    currentSearchIndex = -1

    txtSearchResults.Text = ""
    txtStatusBar.Text = "Search cleared"
  End Sub
  ' ========== Export Image Functionality ==========

  Private Sub btnExportImage_Click(sender As Object, e As RoutedEventArgs)
    If currentDiagramData Is Nothing Then
      MessageBox.Show("No diagram loaded to export.", "Export Error", MessageBoxButton.OK, MessageBoxImage.Warning)
      Return
    End If

    ' Create save file dialog
    Dim saveDialog As New Microsoft.Win32.SaveFileDialog()
    saveDialog.Filter = "PNG Image (*.png)|*.png|JPEG Image (*.jpg)|*.jpg|BMP Image (*.bmp)|*.bmp"
    saveDialog.Title = "Export Diagram as Image"
    saveDialog.FileName = currentDiagramData.metadata.projectName & "_diagram"

    If saveDialog.ShowDialog() = True Then
      Try
        ' Temporarily clear search highlights for clean export
        Dim hadSearchResults As Boolean = searchResults.Count > 0
        If hadSearchResults Then
          ClearSearch()
        End If

        ' Save current transform states
        Dim oldScaleX As Double = scaleTransform.ScaleX
        Dim oldScaleY As Double = scaleTransform.ScaleY
        Dim oldTranslateX As Double = translateTransform.X
        Dim oldTranslateY As Double = translateTransform.Y

        ' Reset transforms for clean export
        scaleTransform.ScaleX = 1.0
        scaleTransform.ScaleY = 1.0
        translateTransform.X = 0
        translateTransform.Y = 0

        ' Force layout update
        canvasContainer.UpdateLayout()

        ' Create render target bitmap
        Dim renderBitmap As New RenderTargetBitmap(
            CInt(diagramCanvas.Width),
            CInt(diagramCanvas.Height),
            96, 96, ' DPI
            PixelFormats.Pbgra32)

        ' Render the canvas
        renderBitmap.Render(diagramCanvas)

        ' Select encoder based on file extension
        Dim encoder As BitmapEncoder
        Dim extension As String = System.IO.Path.GetExtension(saveDialog.FileName).ToLower()

        Select Case extension
          Case ".jpg", ".jpeg"
            Dim jpegEncoder As New JpegBitmapEncoder()
            jpegEncoder.QualityLevel = 95
            encoder = jpegEncoder
          Case ".bmp"
            encoder = New BmpBitmapEncoder()
          Case Else ' Default to PNG
            encoder = New PngBitmapEncoder()
        End Select

        ' Add frame to encoder
        encoder.Frames.Add(BitmapFrame.Create(renderBitmap))

        ' Save to file
        Using stream As New IO.FileStream(saveDialog.FileName, IO.FileMode.Create)
          encoder.Save(stream)
        End Using

        ' Restore transforms
        scaleTransform.ScaleX = oldScaleX
        scaleTransform.ScaleY = oldScaleY
        translateTransform.X = oldTranslateX
        translateTransform.Y = oldTranslateY

        txtStatusBar.Text = "Diagram exported successfully to " & System.IO.Path.GetFileName(saveDialog.FileName)
        MessageBox.Show("Diagram exported successfully!" & vbCrLf & vbCrLf &
                       "Saved to: " & saveDialog.FileName,
                       "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information)

      Catch ex As Exception
        MessageBox.Show("Error exporting diagram: " & ex.Message, "Export Error",
                       MessageBoxButton.OK, MessageBoxImage.Error)
        txtStatusBar.Text = "Export failed"
      End Try
    End If
  End Sub

  ' ========== Open JSON file ==========

  Private Sub btnOpenJson_Click(sender As Object, e As RoutedEventArgs)
    ' Open file dialog (WPF version)
    Dim openDialog As New Microsoft.Win32.OpenFileDialog()
    openDialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
    openDialog.Title = "Select Diagram JSON File"

    If openDialog.ShowDialog() = True Then
      Try
        ' Read JSON file
        Dim jsonText As String = File.ReadAllText(openDialog.FileName)

        ' Parse JSON into our data model
        currentDiagramData = JsonConvert.DeserializeObject(Of DiagramData)(jsonText)

        ' Update UI
        txtStatus.Text = "Loaded: " & System.IO.Path.GetFileName(openDialog.FileName)
        txtStatusBar.Text = "File loaded successfully!"

        ' Display summary information
        Dim summary As String = BuildSummary()
        MessageBox.Show(summary, "JSON Loaded Successfully",
                       MessageBoxButton.OK, MessageBoxImage.Information)

        ' Draw the diagram
        DrawDiagram()

        ' Reset view when loading new diagram
        btnResetView_Click(Nothing, Nothing)

        ' Clear any previous search
        ClearSearch()
        txtSearch.Clear()

      Catch ex As Exception
        MessageBox.Show("Error loading file: " & ex.Message & vbCrLf & vbCrLf &
                       "Stack trace: " & ex.StackTrace,
                       "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        txtStatusBar.Text = "Error loading file"
      End Try
    End If
  End Sub

  Private Function BuildSummary() As String
    ' Build a summary of what was loaded
    If currentDiagramData Is Nothing Then
      Return "No data loaded"
    End If

    Dim summary As String = ""
    summary &= "Project: " & currentDiagramData.metadata.projectName & vbCrLf
    summary &= "Generated: " & currentDiagramData.metadata.generatedDate & vbCrLf
    summary &= "Excel File: " & System.IO.Path.GetFileName(currentDiagramData.metadata.excelFilePath) & vbCrLf
    summary &= vbCrLf
    summary &= "Jobs: " & currentDiagramData.jobs.Count & vbCrLf

    ' Count total datasets
    Dim totalDatasets As Integer = 0
    For Each job In currentDiagramData.jobs
      totalDatasets += job.datasets.Count
    Next
    summary &= "Datasets: " & totalDatasets & vbCrLf
    summary &= vbCrLf

    ' List job names
    summary &= "Job Names:" & vbCrLf
    For Each job In currentDiagramData.jobs
      summary &= "  - " & job.name & " (" & job.datasets.Count & " datasets)" & vbCrLf
    Next

    Return summary
  End Function

  Private Sub DrawDiagram()
    ' Clear the canvas
    diagramCanvas.Children.Clear()

    If currentDiagramData Is Nothing Then
      Return
    End If

    ' Layout constants are now defined at class level

    ' ========== PHASE 1: Analyze shared datasets across all jobs ==========
    ' Dictionary: DatasetName -> List of (JobName, Relationship)
    Dim datasetUsage As New Dictionary(Of String, List(Of Tuple(Of String, String)))

    ' Build the usage map
    For Each job In currentDiagramData.jobs
      For Each dataset In job.datasets
        If Not datasetUsage.ContainsKey(dataset.name) Then
          datasetUsage(dataset.name) = New List(Of Tuple(Of String, String))
        End If
        datasetUsage(dataset.name).Add(New Tuple(Of String, String)(job.name, dataset.relationship))
      Next
    Next

    ' ========== PHASE 2: Categorize and deduplicate datasets ==========
    Dim inputDatasets As New List(Of DatasetInfo)      ' INPUT only (left column)
    Dim outputDatasets As New List(Of DatasetInfo)     ' OUTPUT only (right column)
    Dim bothDatasets As New Dictionary(Of String, List(Of DatasetInfo)) ' BOTH datasets per job (middle column)
    Dim datasetPositions As New Dictionary(Of String, Point)
    Dim processedDatasets As New HashSet(Of String)

    ' Collect all unique datasets and categorize them
    For Each job In currentDiagramData.jobs
      For Each dataset In job.datasets
        If processedDatasets.Contains(dataset.name) Then
          Continue For
        End If
        processedDatasets.Add(dataset.name)

        ' Determine predominant relationship for shared datasets
        Dim relationships = datasetUsage(dataset.name).Select(Function(t) t.Item2).Distinct().ToList()

        ' Categorize based on relationship
        If relationships.Contains("BOTH") OrElse
           (relationships.Contains("INPUT") AndAlso relationships.Contains("OUTPUT")) Then
          ' BOTH datasets go in middle column under their job(s)
          For Each usage In datasetUsage(dataset.name)
            If Not bothDatasets.ContainsKey(usage.Item1) Then
              bothDatasets(usage.Item1) = New List(Of DatasetInfo)
            End If
            ' Only add if not already in this job's BOTH list
            If Not bothDatasets(usage.Item1).Any(Function(d) d.name = dataset.name) Then
              bothDatasets(usage.Item1).Add(dataset)
            End If
          Next
        ElseIf relationships.Contains("OUTPUT") Then
          outputDatasets.Add(dataset)
        Else ' Pure INPUT
          inputDatasets.Add(dataset)
        End If
      Next
    Next

    ' ========== PHASE 3: Calculate column positions ==========
    Dim leftColumnX As Double = START_X
    Dim centerColumnX As Double = START_X + DATASET_WIDTH + COLUMN_GAP
    Dim rightColumnX As Double = centerColumnX + Math.Max(JOB_WIDTH, DATASET_WIDTH) + COLUMN_GAP

    ' Position INPUT datasets (left column) and track the actual bottom Y
    Dim currentY As Double = START_Y
    Dim inputMaxY As Double = START_Y
    For Each dataset In inputDatasets
      datasetPositions(dataset.name) = New Point(leftColumnX, currentY)
      DrawDataset(dataset, leftColumnX, currentY, DATASET_WIDTH, DATASET_HEIGHT)

      ' Use extra spacing for database shapes
      Dim spacing As Double = If(dataset.type = "SQL", SPACING_Y_DATABASE, SPACING_Y)
      currentY += DATASET_HEIGHT + spacing
      inputMaxY = currentY ' Track the actual bottom position
    Next

    ' Position OUTPUT datasets (right column) and track the actual bottom Y
    currentY = START_Y
    Dim outputMaxY As Double = START_Y
    For Each dataset In outputDatasets
      datasetPositions(dataset.name) = New Point(rightColumnX, currentY)
      DrawDataset(dataset, rightColumnX, currentY, DATASET_WIDTH, DATASET_HEIGHT)

      ' Use extra spacing for database shapes
      Dim spacing As Double = If(dataset.type = "SQL", SPACING_Y_DATABASE, SPACING_Y)
      currentY += DATASET_HEIGHT + spacing
      outputMaxY = currentY ' Track the actual bottom position
    Next

    ' ========== PHASE 4: Position jobs and BOTH datasets (center column) ==========
    Dim jobPositions As New Dictionary(Of String, Point)
    Dim jobMaxY As Double = START_Y
    currentY = START_Y

    For Each job In currentDiagramData.jobs
      ' Find all INPUT datasets for this job
      Dim jobInputs = job.datasets.Where(Function(d) d.relationship = "INPUT").ToList()

      ' Calculate vertical center of input datasets
      Dim jobY As Double = currentY

      If jobInputs.Count > 0 Then
        ' Get Y positions of all input datasets
        Dim inputYPositions As New List(Of Double)
        For Each dataset In jobInputs
          If datasetPositions.ContainsKey(dataset.name) Then
            Dim datasetY = datasetPositions(dataset.name).Y
            inputYPositions.Add(datasetY + DATASET_HEIGHT / 2) ' Center of dataset
          End If
        Next

        If inputYPositions.Count > 0 Then
          ' Calculate average Y position (center of all inputs)
          Dim avgY As Double = inputYPositions.Average()
          ' Position job centered at this Y
          jobY = avgY - JOB_HEIGHT / 2
        End If
      End If

      ' Ensure job doesn't overlap with previous job + BOTH datasets
      If jobY < currentY Then
        jobY = currentY
      End If

      jobPositions(job.name) = New Point(centerColumnX, jobY)
      DrawJobRectangle(job.name, centerColumnX, jobY, JOB_WIDTH, JOB_HEIGHT)

      ' Position BOTH datasets for this job (below the job in middle column)
      Dim bothY As Double = jobY + JOB_HEIGHT + JOB_TO_BOTH_GAP

      If bothDatasets.ContainsKey(job.name) Then
        For Each bothDataset In bothDatasets(job.name)
          ' Center BOTH datasets horizontally under the job
          Dim bothX As Double = centerColumnX + (JOB_WIDTH - DATASET_WIDTH) / 2 + BOTH_HORIZONTAL_OFFSET
          If JOB_WIDTH < DATASET_WIDTH Then
            bothX = centerColumnX - (DATASET_WIDTH - JOB_WIDTH) / 2 + BOTH_HORIZONTAL_OFFSET
          End If

          ' Store position AND draw immediately (BEFORE connections)
          datasetPositions(bothDataset.name) = New Point(bothX, bothY)
          DrawDataset(bothDataset, bothX, bothY, DATASET_WIDTH, DATASET_HEIGHT)

          ' Use extra spacing for database shapes
          Dim spacing As Double = If(bothDataset.type = "SQL", SPACING_Y_DATABASE, SPACING_Y)
          bothY += DATASET_HEIGHT + spacing
        Next
      End If

      ' Draw red connection lines from job to BOTH datasets
      If bothDatasets.ContainsKey(job.name) Then
        ' Get job position from the dictionary
        Dim currentJobPos = jobPositions(job.name)

        For Each bothDataset In bothDatasets(job.name)
          Dim datasetPos = datasetPositions(bothDataset.name)

          ' Connection points:
          ' From: Bottom-LEFT of job (left edge, bottom corner)
          Dim twoPixels As Double = 2.0 ' Small offset to avoid overlap with job rectangle border
          Dim jobBottomLeftX As Double = currentJobPos.X + twoPixels
          Dim jobBottomLeftY As Double = currentJobPos.Y + JOB_HEIGHT

          ' To: Left-MIDDLE of dataset (left edge, vertically centered)
          Dim datasetLeftMiddleX As Double = datasetPos.X
          Dim datasetLeftMiddleY As Double = datasetPos.Y + (DATASET_HEIGHT / 2)

          ' Draw the connection with double arrows (red color)
          DrawJobToBothConnection(jobBottomLeftX, jobBottomLeftY,
                                  datasetLeftMiddleX, datasetLeftMiddleY, Brushes.Red)
        Next
      End If

      ' Update currentY for next job (must be below current job's BOTH datasets)
      currentY = Math.Max(jobY + JOB_HEIGHT, bothY) + MIN_JOB_SPACING

      ' Track maximum Y position
      jobMaxY = Math.Max(jobMaxY, currentY)
    Next

    ' ========== PHASE 5: Draw connections with curved arrows (DEDUPLICATED) ==========

    ' Track which dataset-job connections have been drawn to avoid duplicates
    Dim drawnConnections As New HashSet(Of String)

    For Each job In currentDiagramData.jobs
      Dim jobPos = jobPositions(job.name)
      Dim jobCenterY = jobPos.Y + JOB_HEIGHT / 2

      ' Group datasets by name to handle duplicates within same job
      Dim datasetGroups = job.datasets.GroupBy(Function(d) d.name)

      For Each datasetGroup In datasetGroups
        Dim datasetName As String = datasetGroup.Key
        Dim datasetInstances = datasetGroup.ToList()

        ' Create unique key for this connection
        Dim connectionKey As String = datasetName & "|" & job.name

        ' Skip if already drawn
        If drawnConnections.Contains(connectionKey) Then
          Continue For
        End If
        drawnConnections.Add(connectionKey)

        ' Determine the effective relationship for this dataset
        Dim relationships = datasetInstances.Select(Function(d) d.relationship).Distinct().ToList()
        Dim effectiveRelationship As String

        ' If dataset appears multiple times with different relationships, consolidate
        If relationships.Contains("BOTH") OrElse
           (relationships.Contains("INPUT") AndAlso relationships.Contains("OUTPUT")) Then
          effectiveRelationship = "BOTH"
        ElseIf relationships.Contains("OUTPUT") Then
          effectiveRelationship = "OUTPUT"
        Else
          effectiveRelationship = "INPUT"
        End If

        ' Get the first instance for position/type info
        Dim dataset = datasetInstances.First()
        Dim datasetPos = datasetPositions(datasetName)
        Dim datasetCenterY = datasetPos.Y + DATASET_HEIGHT / 2

        ' Draw connection based on effective relationship
        Select Case effectiveRelationship
          Case "INPUT"
            ' Curved arrow from dataset (left) to job (center)
            Dim startX = datasetPos.X + DATASET_WIDTH
            Dim startY = datasetCenterY
            Dim endX = jobPos.X
            Dim endY = jobCenterY
            DrawCurvedArrow(startX, startY, endX, endY, Brushes.Blue, curveToRight:=True)

          Case "OUTPUT"
            ' Curved arrow from job (center) to dataset (right)
            Dim startX = jobPos.X + JOB_WIDTH
            Dim startY = jobCenterY
            Dim endX = datasetPos.X
            Dim endY = datasetCenterY
            DrawCurvedArrow(startX, startY, endX, endY, Brushes.Green, curveToRight:=True)

          Case "BOTH"
            ' Skip drawing connections for BOTH datasets
            ' The datasets are already drawn below the job in the center column
            Continue For
        End Select
      Next
    Next

    ' ========== PHASE 6: Auto-resize canvas to fit all content ==========
    ' Calculate the actual maximum Y position from all three columns
    Dim maxY As Double = Math.Max(inputMaxY, Math.Max(outputMaxY, jobMaxY))

    ' Add padding at the bottom
    Dim canvasWidth As Double = rightColumnX + DATASET_WIDTH + 50
    Dim canvasHeight As Double = maxY + 50
    diagramCanvas.Width = canvasWidth
    diagramCanvas.Height = canvasHeight

    ' Count shared datasets
    Dim sharedCount As Integer = datasetUsage.Where(Function(kvp) kvp.Value.Count > 1).Count()
    Dim bothCount As Integer = bothDatasets.Values.Sum(Function(list) list.Count)

    txtStatusBar.Text = $"Diagram rendered: {currentDiagramData.jobs.Count} jobs, " +
                        $"{processedDatasets.Count} unique datasets ({sharedCount} shared, {bothCount} BOTH)"

    ' Enable export button now that diagram is loaded
    btnExportImage.IsEnabled = True
  End Sub

  ' Method to draw red double-arrow line from job to BOTH dataset
  Private Sub DrawJobToBothConnection(jobX As Double, jobY As Double,
                                       datasetX As Double, datasetY As Double,
                                       color As Brush)
    ' Draw straight line with double arrows from job bottom-center to dataset left-middle

    ' Draw the main line
    Dim connectionLine As New Line()
    connectionLine.X1 = jobX
    connectionLine.Y1 = jobY
    connectionLine.X2 = datasetX
    connectionLine.Y2 = datasetY
    connectionLine.Stroke = color
    connectionLine.StrokeThickness = 2
    diagramCanvas.Children.Add(connectionLine)

    ' Calculate angle for arrowheads
    Dim dx As Double = datasetX - jobX
    Dim dy As Double = datasetY - jobY
    Dim angle As Double = Math.Atan2(dy, dx) * 180 / Math.PI

    ' Arrow at dataset end (pointing INTO dataset)
    Dim arrowToDataset As New Polygon()
    arrowToDataset.Fill = color
    arrowToDataset.Points = New PointCollection From {
        New Point(0, 0),
        New Point(-10, -5),
        New Point(-10, 5)
    }

    Dim transformToDataset As New TransformGroup()
    transformToDataset.Children.Add(New RotateTransform(angle))
    transformToDataset.Children.Add(New TranslateTransform(datasetX, datasetY))
    arrowToDataset.RenderTransform = transformToDataset
    diagramCanvas.Children.Add(arrowToDataset)

    ' Arrow at job end (pointing DOWN from job)
    Dim arrowFromJob As New Polygon()
    arrowFromJob.Fill = color
    arrowFromJob.Points = New PointCollection From {
        New Point(0, 0),
        New Point(-10, -5),
        New Point(-10, 5)
    }

    Dim transformFromJob As New TransformGroup()
    transformFromJob.Children.Add(New RotateTransform(angle + 180)) ' Reverse direction
    transformFromJob.Children.Add(New TranslateTransform(jobX, jobY))
    arrowFromJob.RenderTransform = transformFromJob
    diagramCanvas.Children.Add(arrowFromJob)
  End Sub

  ' New method to draw vertical double arrows for BOTH relationships with zigzag routing
  Private Sub DrawVerticalDoubleArrow(x1 As Double, y1 As Double, x2 As Double, y2 As Double, color As Brush)
    ' Draw a zigzag path from job to BOTH dataset (wiring diagram style)
    ' x1, y1 = bottom center of job
    ' x2, y2 = top-left corner of dataset

    ' Calculate connection points
    Dim jobBottomX As Double = x1
    Dim jobBottomY As Double = y1

    ' Calculate the LEFT side of the dataset (for connection point)
    Dim datasetLeftX As Double = x2 - (DATASET_WIDTH / 2)
    Dim datasetCenterY As Double = y2 + (DATASET_HEIGHT / 2)

    ' Define zigzag path points
    Const INITIAL_DROP As Double = 10    ' Initial drop from job
    Const LEFT_OFFSET As Double = 50     ' How far left to go from job center
    Const LEFT_CLEARANCE As Double = 25  ' Distance to the LEFT of dataset edge

    ' The vertical routing line X position (to the left of job center)
    Dim routingLineX As Double = jobBottomX - LEFT_OFFSET

    ' The horizontal connection endpoint (to the left of dataset)
    Dim connectionX As Double = datasetLeftX - LEFT_CLEARANCE

    ' ===== Path breakdown =====
    ' Point 1: Start at job bottom center
    Dim p1X As Double = jobBottomX
    Dim p1Y As Double = jobBottomY

    ' Point 2: Drop down from job
    Dim p2X As Double = p1X
    Dim p2Y As Double = p1Y + INITIAL_DROP

    ' Point 3: Go left to routing line
    Dim p3X As Double = routingLineX
    Dim p3Y As Double = p2Y

    ' Point 4: Go down on routing line to dataset center height
    Dim p4X As Double = routingLineX
    Dim p4Y As Double = datasetCenterY

    ' Point 5: Go right to connection point (left of dataset)
    Dim p5X As Double = connectionX
    Dim p5Y As Double = datasetCenterY

    ' Draw the complete path segments
    ' Segment 1: Job center down to turning point
    Dim line1 As New Line()
    line1.X1 = p1X
    line1.Y1 = p1Y
    line1.X2 = p2X
    line1.Y2 = p2Y
    line1.Stroke = color
    line1.StrokeThickness = 2
    diagramCanvas.Children.Add(line1)

    ' Segment 2: Turn left to routing line
    Dim line2 As New Line()
    line2.X1 = p2X
    line2.Y1 = p2Y
    line2.X2 = p3X
    line2.Y2 = p3Y
    line2.Stroke = color
    line2.StrokeThickness = 2
    diagramCanvas.Children.Add(line2)

    ' Segment 3: Go down routing line to dataset center height
    Dim line3 As New Line()
    line3.X1 = p3X
    line3.Y1 = p3Y
    line3.X2 = p4X
    line3.Y2 = p4Y
    line3.Stroke = color
    line3.StrokeThickness = 2
    diagramCanvas.Children.Add(line3)

    ' Segment 4: Go right to dataset
    Dim line4 As New Line()
    line4.X1 = p4X
    line4.Y1 = p4Y
    line4.X2 = p5X
    line4.Y2 = p5Y
    line4.Stroke = color
    line4.StrokeThickness = 2
    diagramCanvas.Children.Add(line4)

    ' Arrowhead pointing right INTO dataset
    Dim arrowToDataset As New Polygon()
    arrowToDataset.Fill = color
    arrowToDataset.Points = New PointCollection From {
        New Point(p5X, p5Y),
        New Point(p5X - 10, p5Y - 5),
        New Point(p5X - 10, p5Y + 5)
    }
    diagramCanvas.Children.Add(arrowToDataset)

    ' Arrowhead pointing DOWN from job (only draw on first connection)
    ' To avoid multiple arrows at job bottom, we'll always draw it
    ' (The calling code should handle drawing this only once per job if needed)
    Dim arrowFromJob As New Polygon()
    arrowFromJob.Fill = color
    arrowFromJob.Points = New PointCollection From {
        New Point(p1X, p1Y),
        New Point(p1X - 5, p1Y + 10),
        New Point(p1X + 5, p1Y + 10)
    }
    diagramCanvas.Children.Add(arrowFromJob)
  End Sub

  Private Sub DrawCurvedArrow(x1 As Double, y1 As Double, x2 As Double, y2 As Double,
                              color As Brush, Optional curveToRight As Boolean = True)
    ' Draw a smooth Bézier curve arrow between two points

    ' Calculate control points for the curve
    Dim dx As Double = x2 - x1
    Dim dy As Double = y2 - y1
    Dim distance As Double = Math.Sqrt(dx * dx + dy * dy)

    ' Curve intensity based on distance
    Dim curveAmount As Double = Math.Min(distance * 0.3, 60)

    ' Calculate perpendicular offset for curve direction
    Dim perpX As Double = -dy / distance * curveAmount
    Dim perpY As Double = dx / distance * curveAmount

    If Not curveToRight Then
      perpX = -perpX
      perpY = -perpY
    End If

    ' Control points for quadratic Bézier curve
    Dim midX As Double = (x1 + x2) / 2
    Dim midY As Double = (y1 + y2) / 2
    Dim controlX As Double = midX + perpX
    Dim controlY As Double = midY + perpY

    ' Create path geometry
    Dim pathFigure As New PathFigure()
    pathFigure.StartPoint = New Point(x1, y1)

    ' Quadratic Bézier segment
    Dim bezierSegment As New QuadraticBezierSegment()
    bezierSegment.Point1 = New Point(controlX, controlY)
    bezierSegment.Point2 = New Point(x2, y2)
    pathFigure.Segments.Add(bezierSegment)

    Dim pathGeometry As New PathGeometry()
    pathGeometry.Figures.Add(pathFigure)

    ' FIX: Fully qualify Path to avoid ambiguity with System.IO.Path
    Dim path As New System.Windows.Shapes.Path()
    path.Stroke = color
    path.StrokeThickness = 2
    path.Data = pathGeometry
    diagramCanvas.Children.Add(path)

    ' Draw arrowhead at the end point
    DrawArrowhead(x2, y2, controlX, controlY, color)
  End Sub

  Private Sub DrawArrowhead(x As Double, y As Double, controlX As Double, controlY As Double, color As Brush)
    ' Draw arrowhead pointing in the direction of the curve

    ' Calculate angle from control point to end point
    Dim dx As Double = x - controlX
    Dim dy As Double = y - controlY
    Dim angle As Double = Math.Atan2(dy, dx) * 180 / Math.PI

    ' Create arrowhead polygon
    Dim arrowhead As New Polygon()
    arrowhead.Fill = color
    arrowhead.Points = New PointCollection From {
        New Point(0, 0),
        New Point(-10, -5),
        New Point(-10, 5)
    }

    ' Transform arrowhead to the correct position and angle
    Dim transform As New TransformGroup()
    transform.Children.Add(New RotateTransform(angle))
    transform.Children.Add(New TranslateTransform(x, y))
    arrowhead.RenderTransform = transform

    diagramCanvas.Children.Add(arrowhead)
  End Sub

  Private Sub DrawCurvedDoubleArrow(x1 As Double, y1 As Double, x2 As Double, y2 As Double, color As Brush)
    ' Draw curved double arrow for BOTH relationships with arrowheads on both ends

    Dim dx As Double = x2 - x1
    Dim dy As Double = y2 - y1
    Dim distance As Double = Math.Sqrt(dx * dx + dy * dy)

    ' Curve intensity based on distance
    Dim curveAmount As Double = Math.Min(distance * 0.3, 60)

    ' Calculate perpendicular offset for curve direction
    Dim perpX As Double = -dy / distance * curveAmount
    Dim perpY As Double = dx / distance * curveAmount

    ' Control points for quadratic Bézier curve
    Dim midX As Double = (x1 + x2) / 2
    Dim midY As Double = (y1 + y2) / 2
    Dim controlX As Double = midX + perpX
    Dim controlY As Double = midY + perpY

    ' Create path geometry for the curve
    Dim pathFigure As New PathFigure()
    pathFigure.StartPoint = New Point(x1, y1)

    ' Quadratic Bézier segment
    Dim bezierSegment As New QuadraticBezierSegment()
    bezierSegment.Point1 = New Point(controlX, controlY)
    bezierSegment.Point2 = New Point(x2, y2)
    pathFigure.Segments.Add(bezierSegment)

    Dim pathGeometry As New PathGeometry()
    pathGeometry.Figures.Add(pathFigure)

    ' Draw the single curved line
    Dim path As New System.Windows.Shapes.Path()
    path.Stroke = color
    path.StrokeThickness = 2
    path.Data = pathGeometry
    diagramCanvas.Children.Add(path)

    ' Draw arrowhead at the end point (x2, y2) - pointing toward dataset
    DrawArrowhead(x2, y2, controlX, controlY, color)

    ' Draw arrowhead at the start point (x1, y1) - pointing toward job
    ' Calculate angle from control point to start point (reverse direction)
    Dim dxStart As Double = x1 - controlX
    Dim dyStart As Double = y1 - controlY
    Dim angleStart As Double = Math.Atan2(dyStart, dxStart) * 180 / Math.PI

    ' Create arrowhead polygon for start point
    Dim arrowheadStart As New Polygon()
    arrowheadStart.Fill = color
    arrowheadStart.Points = New PointCollection From {
        New Point(0, 0),
        New Point(-10, -5),
        New Point(-10, 5)
    }

    ' Transform arrowhead to the correct position and angle
    Dim transformStart As New TransformGroup()
    transformStart.Children.Add(New RotateTransform(angleStart))
    transformStart.Children.Add(New TranslateTransform(x1, y1))
    arrowheadStart.RenderTransform = transformStart

    diagramCanvas.Children.Add(arrowheadStart)
  End Sub

  Private Function WrapDatasetNameAtQualifier(datasetName As String, maxWidth As Double, fontSize As Double) As String
    ' Wrap dataset name at qualifier boundaries (periods) for mainframe datasets
    ' Rules:
    ' - Qualifiers are separated by periods
    ' - Wrap should occur at a period when text exceeds maxWidth
    ' - Period stays at end of first line
    ' - Second line starts with next character (not the period)

    ' Create a temporary TextBlock to measure text width
    Dim measureBlock As New TextBlock()
    measureBlock.FontSize = fontSize
    measureBlock.Text = datasetName
    measureBlock.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))

    ' If the text fits, no wrapping needed
    If measureBlock.DesiredSize.Width <= maxWidth Then
      Return datasetName
    End If

    ' Split into qualifiers
    Dim qualifiers As String() = datasetName.Split("."c)

    If qualifiers.Length <= 1 Then
      ' No periods to wrap at, return as-is
      Return datasetName
    End If

    ' Find the optimal wrap point
    Dim bestWrapIndex As Integer = -1
    Dim firstLine As String = ""

    ' Try each qualifier position to find best wrap point
    For i As Integer = 0 To qualifiers.Length - 2
      ' Build first line with period at end
      firstLine = String.Join(".", qualifiers.Take(i + 1)) & "."

      ' Measure first line
      measureBlock.Text = firstLine
      measureBlock.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))

      ' If first line fits, this is a candidate
      If measureBlock.DesiredSize.Width <= maxWidth Then
        bestWrapIndex = i
      Else
        ' First line too long, use previous wrap point
        Exit For
      End If
    Next

    ' If we found a valid wrap point, apply it
    If bestWrapIndex >= 0 AndAlso bestWrapIndex < qualifiers.Length - 1 Then
      firstLine = String.Join(".", qualifiers.Take(bestWrapIndex + 1)) & "."
      Dim secondLine As String = String.Join(".", qualifiers.Skip(bestWrapIndex + 1))
      Return firstLine & vbCrLf & secondLine
    End If

    ' Fallback: wrap at first period if nothing else works
    If qualifiers.Length >= 2 Then
      Return qualifiers(0) & "." & vbCrLf & String.Join(".", qualifiers.Skip(1))
    End If

    ' Last resort: return original name
    Return datasetName
  End Function

  Private Sub DrawJobRectangle(jobName As String, x As Double, y As Double, width As Double, height As Double)
    ' Create rectangle for job
    Dim rect As New Rectangle()
    rect.Width = width
    rect.Height = height
    rect.Fill = Brushes.LightSteelBlue
    rect.Stroke = Brushes.DarkBlue
    rect.StrokeThickness = 1
    Canvas.SetLeft(rect, x)
    Canvas.SetTop(rect, y)
    diagramCanvas.Children.Add(rect)

    ' Add job name label
    Dim label As New TextBlock()
    label.Text = jobName
    label.FontSize = 12
    label.FontWeight = FontWeights.Bold
    label.Foreground = Brushes.DarkBlue
    label.TextAlignment = TextAlignment.Center
    label.Width = width
    Canvas.SetLeft(label, x)
    Canvas.SetTop(label, y + (height / 2) - 8)
    diagramCanvas.Children.Add(label)
  End Sub

  Private Sub DrawDataset(dataset As DatasetInfo, x As Double, y As Double, width As Double, height As Double)
    ' Determine color based on dataset type
    Dim fillBrush As Brush = Brushes.White
    Select Case dataset.type
      Case "GDG"
        fillBrush = Brushes.LightGreen
      Case "PDS"
        fillBrush = Brushes.LightCyan
      Case "Library"
        fillBrush = Brushes.LightBlue
      Case "File"
        fillBrush = Brushes.LightSalmon
      Case "SQL"
        fillBrush = Brushes.LightYellow
      Case Else
        fillBrush = Brushes.White
    End Select

    ' Create shape based on dataset type (PlantUML style)
    Select Case dataset.type
      Case "SQL"
        ' Database shape (cylinder) - already implemented
        DrawDatabaseShape(dataset, x, y, width, height, fillBrush)
        Return

      Case "Library"
        ' Folder shape (folder with tab)
        DrawFolderShape(dataset, x, y, width, height, fillBrush)
        Return

      Case "PDS"
        ' Package shape (3D box)
        DrawPackageShape(dataset, x, y, width, height, fillBrush)
        Return

      Case "GDG"
        ' Collections shape (framed box)
        DrawCollectionsShape(dataset, x, y, width, height, fillBrush)
        Return

      Case "File"
        ' Queue shape (rounded rectangle with lines)
        DrawQueueShape(dataset, x, y, width, height, fillBrush)
        Return

      Case Else
        ' Default: regular rectangle for unknown types
        Dim rect = New Rectangle()
        rect.Width = width
        rect.Height = height
        rect.Fill = fillBrush
        rect.Stroke = Brushes.Gray
        rect.StrokeThickness = 1.5
        rect.Cursor = Cursors.Hand
        rect.Tag = dataset
        Panel.SetZIndex(rect, 1) ' BRING TO FRONT

        AddHandler rect.MouseRightButtonDown, AddressOf Dataset_RightClick

        Canvas.SetLeft(rect, x)
        Canvas.SetTop(rect, y)
        diagramCanvas.Children.Add(rect)

        ' Add dataset name label
        Dim label As New TextBlock()
        label.Text = dataset.name & vbCrLf & "(" & dataset.excelReference.worksheet & ":" & dataset.excelReference.cellReference & ")"
        label.FontSize = 9
        label.Foreground = Brushes.Black
        label.TextAlignment = TextAlignment.Center
        label.Width = width - 10
        label.TextWrapping = TextWrapping.Wrap
        label.IsHitTestVisible = False
        Panel.SetZIndex(label, 2) ' BRING TO FRONT (above shapes)
        Canvas.SetLeft(label, x + 5)
        Canvas.SetTop(label, y + 3)
        diagramCanvas.Children.Add(label)
    End Select
  End Sub
  Private Sub DrawQueueShape(dataset As DatasetInfo, x As Double, y As Double, width As Double, height As Double, fillBrush As Brush)
    ' Create a queue shape (capsule/rounded rectangle - PlantUML style)
    ' Queue should look like a horizontal capsule with rounded ends

    Dim cornerRadius As Double = height / 2 ' Make ends fully rounded (capsule shape)

    ' Main rounded rectangle (capsule)
    Dim rect As New Rectangle()
    rect.Width = width
    rect.Height = height
    rect.Fill = fillBrush
    rect.Stroke = Brushes.Gray
    rect.StrokeThickness = 1.5
    rect.RadiusX = cornerRadius
    rect.RadiusY = cornerRadius
    rect.Cursor = Cursors.Hand
    rect.Tag = dataset
    Panel.SetZIndex(rect, 1) ' BRING TO FRONT

    AddHandler rect.MouseRightButtonDown, AddressOf Dataset_RightClick

    Canvas.SetLeft(rect, x)
    Canvas.SetTop(rect, y)
    diagramCanvas.Children.Add(rect)

    ' Add dataset name label with intelligent wrapping at qualifiers
    Dim label As New TextBlock()

    ' Apply intelligent wrapping for mainframe dataset names
    Dim wrappedName As String = WrapDatasetNameAtQualifier(dataset.name, width - 10, 9)

    label.Text = wrappedName & vbCrLf & "(" & dataset.excelReference.worksheet & ":" & dataset.excelReference.cellReference & ")"
    label.FontSize = 9
    label.Foreground = Brushes.Black
    label.TextAlignment = TextAlignment.Center
    label.Width = width - 10
    label.TextWrapping = TextWrapping.NoWrap ' Manual wrapping at qualifiers
    label.IsHitTestVisible = False
    Panel.SetZIndex(label, 2) ' BRING TO FRONT (above shapes)
    Canvas.SetLeft(label, x + 5)
    Canvas.SetTop(label, y + (height / 2) - 10) ' Center vertically
    diagramCanvas.Children.Add(label)
  End Sub
  Private Sub DrawFolderShape(dataset As DatasetInfo, x As Double, y As Double, width As Double, height As Double, fillBrush As Brush)
    ' Create a folder shape (PlantUML style)
    Dim tabWidth As Double = width * 0.4
    Dim tabHeight As Double = height * 0.25

    ' Create the folder body path
    Dim pathFigure As New PathFigure()
    pathFigure.StartPoint = New Point(x, y + tabHeight)
    pathFigure.IsClosed = True

    ' Draw folder: tab then main body
    pathFigure.Segments.Add(New LineSegment(New Point(x, y), True)) ' Left edge of tab
    pathFigure.Segments.Add(New LineSegment(New Point(x + tabWidth, y), True)) ' Top of tab
    pathFigure.Segments.Add(New LineSegment(New Point(x + tabWidth + 10, y + tabHeight), True)) ' Tab slope
    pathFigure.Segments.Add(New LineSegment(New Point(x + width, y + tabHeight), True)) ' Top right
    pathFigure.Segments.Add(New LineSegment(New Point(x + width, y + height), True)) ' Right edge
    pathFigure.Segments.Add(New LineSegment(New Point(x, y + height), True)) ' Bottom edge
    pathFigure.Segments.Add(New LineSegment(New Point(x, y + tabHeight), True)) ' Left edge

    Dim pathGeometry As New PathGeometry()
    pathGeometry.Figures.Add(pathFigure)

    Dim path As New System.Windows.Shapes.Path()
    path.Data = pathGeometry
    path.Fill = fillBrush
    path.Stroke = Brushes.Gray
    path.StrokeThickness = 1.5
    path.Cursor = Cursors.Hand
    path.Tag = dataset
    Panel.SetZIndex(path, 1) ' BRING TO FRONT

    AddHandler path.MouseRightButtonDown, AddressOf Dataset_RightClick
    diagramCanvas.Children.Add(path)

    ' Add dataset name label
    Dim label As New TextBlock()
    label.Text = dataset.name & vbCrLf & "(" & dataset.excelReference.worksheet & ":" & dataset.excelReference.cellReference & ")"
    label.FontSize = 9
    label.Foreground = Brushes.Black
    label.TextAlignment = TextAlignment.Center
    label.Width = width - 10
    label.TextWrapping = TextWrapping.Wrap
    label.IsHitTestVisible = False
    Panel.SetZIndex(label, 2) ' BRING TO FRONT (above shapes)
    Canvas.SetLeft(label, x + 5)
    Canvas.SetTop(label, y + tabHeight + 5)
    diagramCanvas.Children.Add(label)
  End Sub

  Private Sub DrawPackageShape(dataset As DatasetInfo, x As Double, y As Double, width As Double, height As Double, fillBrush As Brush)
    ' PDS now uses simple Rectangle - LightCyan color
    Dim rect As New Rectangle()
    rect.Width = width
    rect.Height = height
    rect.Fill = fillBrush
    rect.Stroke = Brushes.Gray
    rect.StrokeThickness = 1.5
    rect.Cursor = Cursors.Hand
    rect.Tag = dataset
    Panel.SetZIndex(rect, 1) ' BRING TO FRONT

    AddHandler rect.MouseRightButtonDown, AddressOf Dataset_RightClick

    Canvas.SetLeft(rect, x)
    Canvas.SetTop(rect, y)
    diagramCanvas.Children.Add(rect)

    ' Add dataset name label
    Dim label As New TextBlock()
    label.Text = dataset.name & vbCrLf & "(" & dataset.excelReference.worksheet & ":" & dataset.excelReference.cellReference & ")"
    label.FontSize = 9
    label.Foreground = Brushes.Black
    label.TextAlignment = TextAlignment.Center
    label.Width = width - 10
    label.TextWrapping = TextWrapping.Wrap
    label.IsHitTestVisible = False
    Panel.SetZIndex(label, 2) ' BRING TO FRONT (above shapes)
    Canvas.SetLeft(label, x + 5)
    Canvas.SetTop(label, y + (height / 2) - 10) ' Center vertically
    diagramCanvas.Children.Add(label)
  End Sub

  Private Sub DrawCollectionsShape(dataset As DatasetInfo, x As Double, y As Double, width As Double, height As Double, fillBrush As Brush)
    ' GDG now uses Queue/Capsule shape (single rounded Rectangle) - LightGreen color
    Dim cornerRadius As Double = height / 2 ' Make ends fully rounded (capsule shape)

    ' Main rounded rectangle (capsule)
    Dim rect As New Rectangle()
    rect.Width = width
    rect.Height = height
    rect.Fill = fillBrush
    rect.Stroke = Brushes.Gray
    rect.StrokeThickness = 1.5
    rect.RadiusX = cornerRadius
    rect.RadiusY = cornerRadius
    rect.Cursor = Cursors.Hand
    rect.Tag = dataset
    Panel.SetZIndex(rect, 1) ' BRING TO FRONT

    AddHandler rect.MouseRightButtonDown, AddressOf Dataset_RightClick

    Canvas.SetLeft(rect, x)
    Canvas.SetTop(rect, y)
    diagramCanvas.Children.Add(rect)

    ' Add dataset name label
    Dim label As New TextBlock()
    label.Text = dataset.name & vbCrLf & "(" & dataset.excelReference.worksheet & ":" & dataset.excelReference.cellReference & ")"
    label.FontSize = 9
    label.Foreground = Brushes.Black
    label.TextAlignment = TextAlignment.Center
    label.Width = width - 10
    label.TextWrapping = TextWrapping.Wrap
    label.IsHitTestVisible = False
    Panel.SetZIndex(label, 2) ' BRING TO FRONT (above shapes)
    Canvas.SetLeft(label, x + 5)
    Canvas.SetTop(label, y + (height / 2) - 10) ' Center vertically
    diagramCanvas.Children.Add(label)
  End Sub

  Private Sub DrawDatabaseShape(dataset As DatasetInfo, x As Double, y As Double, width As Double, height As Double, fillBrush As Brush)
    ' SQL now uses Folder shape (single Path) - LightYellow color
    Dim tabWidth As Double = width * 0.4
    Dim tabHeight As Double = height * 0.25

    ' Create the folder body path
    Dim pathFigure As New PathFigure()
    pathFigure.StartPoint = New Point(x, y + tabHeight)
    pathFigure.IsClosed = True

    ' Draw folder: tab then main body
    pathFigure.Segments.Add(New LineSegment(New Point(x, y), True)) ' Left edge of tab
    pathFigure.Segments.Add(New LineSegment(New Point(x + tabWidth, y), True)) ' Top of tab
    pathFigure.Segments.Add(New LineSegment(New Point(x + tabWidth + 10, y + tabHeight), True)) ' Tab slope
    pathFigure.Segments.Add(New LineSegment(New Point(x + width, y + tabHeight), True)) ' Top right
    pathFigure.Segments.Add(New LineSegment(New Point(x + width, y + height), True)) ' Right edge
    pathFigure.Segments.Add(New LineSegment(New Point(x, y + height), True)) ' Bottom edge
    pathFigure.Segments.Add(New LineSegment(New Point(x, y + tabHeight), True)) ' Left edge

    Dim pathGeometry As New PathGeometry()
    pathGeometry.Figures.Add(pathFigure)

    Dim path As New System.Windows.Shapes.Path()
    path.Data = pathGeometry
    path.Fill = fillBrush
    path.Stroke = Brushes.Gray
    path.StrokeThickness = 1.5
    path.Cursor = Cursors.Hand
    path.Tag = dataset
    Panel.SetZIndex(path, 1) ' BRING TO FRONT

    AddHandler path.MouseRightButtonDown, AddressOf Dataset_RightClick
    diagramCanvas.Children.Add(path)

    ' Parse dataset name to extract DATABASE.TABLE
    Dim databaseName As String = ""
    Dim tableName As String = dataset.name

    ' Check if dataset name contains a period (DATABASE.TABLE pattern)
    Dim dotIndex As Integer = dataset.name.IndexOf("."c)
    If dotIndex > 0 Then
      ' Split into DATABASE and TABLE
      databaseName = dataset.name.Substring(0, dotIndex)
      tableName = dataset.name.Substring(dotIndex + 1)
    End If

    ' Add database name label on the tab (only if database name exists)
    If Not String.IsNullOrEmpty(databaseName) Then
      Dim tabLabel As New TextBlock()
      tabLabel.Text = databaseName
      tabLabel.FontSize = 8
      tabLabel.FontWeight = FontWeights.Bold
      tabLabel.Foreground = Brushes.DarkGray
      tabLabel.TextAlignment = TextAlignment.Left
      tabLabel.Width = tabWidth
      tabLabel.IsHitTestVisible = False
      Panel.SetZIndex(tabLabel, 2) ' BRING TO FRONT (above shapes)
      Canvas.SetLeft(tabLabel, x + 2) ' Small left padding
      Canvas.SetTop(tabLabel, y + 2) ' Position near top of tab
      diagramCanvas.Children.Add(tabLabel)
    End If

    ' Add table name label with Excel reference
    Dim label As New TextBlock()
    label.Text = tableName & vbCrLf & "(" & dataset.excelReference.worksheet & ":" & dataset.excelReference.cellReference & ")"
    label.FontSize = 9
    label.Foreground = Brushes.Black
    label.TextAlignment = TextAlignment.Center
    label.Width = width - 10
    label.TextWrapping = TextWrapping.Wrap
    label.IsHitTestVisible = False
    Panel.SetZIndex(label, 2) ' BRING TO FRONT (above shapes)
    Canvas.SetLeft(label, x + 5)
    Canvas.SetTop(label, y + tabHeight + 5)
    diagramCanvas.Children.Add(label)
  End Sub

  ' ========== Dataset Right-Click Context Menu ==========

  Private Sub Dataset_RightClick(sender As Object, e As MouseButtonEventArgs)
    ' Get the clicked rectangle
    Dim rect As Rectangle = CType(sender, Rectangle)

    ' Get the dataset info from Tag
    Dim dataset As DatasetInfo = CType(rect.Tag, DatasetInfo)

    If dataset Is Nothing Then
      Return
    End If

    ' Create context menu
    Dim contextMenu As New ContextMenu()

    ' Menu Item 1: Open Excel Cell
    Dim menuOpenExcel As New MenuItem()
    menuOpenExcel.Header = "Open Excel Cell"
    menuOpenExcel.Icon = New TextBlock() With {
      .Text = "📊",
      .FontSize = 14
    }
    AddHandler menuOpenExcel.Click, Sub(s, args) OpenExcelCell(dataset)
    contextMenu.Items.Add(menuOpenExcel)

    ' Separator
    contextMenu.Items.Add(New Separator())

    ' Menu Item 2: Copy Dataset Name
    Dim menuCopyName As New MenuItem()
    menuCopyName.Header = "Copy Dataset Name"
    menuCopyName.Icon = New TextBlock() With {
      .Text = "📋",
      .FontSize = 14
    }
    AddHandler menuCopyName.Click, Sub(s, args) CopyDatasetName(dataset)
    contextMenu.Items.Add(menuCopyName)

    ' Menu Item 3: Copy Diagram Content
    Dim menuCopyContent As New MenuItem()
    menuCopyContent.Header = "Copy Dataset Details"
    menuCopyContent.Icon = New TextBlock() With {
      .Text = "📄",
      .FontSize = 14
    }
    AddHandler menuCopyContent.Click, Sub(s, args) CopyDatasetDetails(dataset)
    contextMenu.Items.Add(menuCopyContent)

    ' Show the context menu at cursor position
    contextMenu.IsOpen = True
    contextMenu.PlacementTarget = rect
    contextMenu.Placement = Primitives.PlacementMode.Mouse

    e.Handled = True
  End Sub

  Private Sub OpenExcelCell(dataset As DatasetInfo)
    ' Show feedback
    txtStatusBar.Text = "Opening Excel: " & dataset.name & " at " & dataset.excelReference.cellReference

    ' Get Excel file path from metadata
    Dim excelFilePath As String = currentDiagramData.metadata.excelFilePath

    ' Check if file exists
    If Not File.Exists(excelFilePath) Then
      ' File not found - prompt user to locate it
      Dim result = MessageBox.Show(
          "Excel file not found at:" & vbCrLf & vbCrLf &
          excelFilePath & vbCrLf & vbCrLf &
          "Would you like to browse for the Excel file?",
          "File Not Found",
          MessageBoxButton.YesNo,
          MessageBoxImage.Question)

      If result = MessageBoxResult.Yes Then
        ' Let user browse for the file
        Dim openDialog As New Microsoft.Win32.OpenFileDialog()
        openDialog.Filter = "Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All Files (*.*)|*.*"
        openDialog.Title = "Locate Excel Model File"
        openDialog.FileName = System.IO.Path.GetFileName(excelFilePath)

        If openDialog.ShowDialog() = True Then
          ' Update the path for this session
          excelFilePath = openDialog.FileName
          currentDiagramData.metadata.excelFilePath = excelFilePath
        Else
          txtStatusBar.Text = "Excel navigation cancelled"
          Return
        End If
      Else
        txtStatusBar.Text = "Excel file not found"
        Return
      End If
    End If

    ' Navigate to Excel
    Try
      ExcelNavigator.NavigateToCell(
          excelFilePath,
          dataset.excelReference.worksheet,
          dataset.excelReference.cellReference
      )
      txtStatusBar.Text = "Opened Excel at " & dataset.excelReference.worksheet & ":" & dataset.excelReference.cellReference
    Catch ex As Exception
      MessageBox.Show("Could not open Excel: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
      txtStatusBar.Text = "Error opening Excel"
    End Try
  End Sub

  Private Sub CopyDatasetName(dataset As DatasetInfo)
    Try
      ' Copy just the dataset name to clipboard
      Clipboard.SetText(dataset.name)
      txtStatusBar.Text = $"Copied dataset name: {dataset.name}"
    Catch ex As Exception
      MessageBox.Show("Failed to copy to clipboard: " & ex.Message, "Copy Error",
                     MessageBoxButton.OK, MessageBoxImage.Error)
      txtStatusBar.Text = "Copy failed"
    End Try
  End Sub

  Private Sub CopyDatasetDetails(dataset As DatasetInfo)
    Try
      ' Build formatted text with all dataset details
      Dim details As New System.Text.StringBuilder()
      details.AppendLine($"Dataset Name: {dataset.name}")
      details.AppendLine($"Type: {dataset.type}")
      details.AppendLine($"Relationship: {dataset.relationship}")
      details.AppendLine($"Worksheet: {dataset.excelReference.worksheet}")
      details.AppendLine($"Cell: {dataset.excelReference.cellReference}")
      details.AppendLine($"Row: {dataset.excelReference.row}")
      details.AppendLine($"Column: {dataset.excelReference.column}")

      ' Copy to clipboard
      Clipboard.SetText(details.ToString())
      txtStatusBar.Text = $"Copied details for: {dataset.name}"
    Catch ex As Exception
      MessageBox.Show("Failed to copy to clipboard: " & ex.Message, "Copy Error",
                     MessageBoxButton.OK, MessageBoxImage.Error)
      txtStatusBar.Text = "Copy failed"
    End Try
  End Sub

  ' Remove or comment out the old Dataset_Click method since we're using right-click now
  ' Private Sub Dataset_Click(sender As Object, e As MouseButtonEventArgs)
  '   ... (no longer needed)
  ' End Sub
End Class