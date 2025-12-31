Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel

Public Class ExcelNavigator
  ''' <summary>
  ''' Opens Excel and navigates to a specific cell, reusing existing Excel instance if available
  ''' </summary>
  Public Shared Sub NavigateToCell(excelFilePath As String, worksheetName As String, cellReference As String)
    Dim excelApp As Excel.Application = Nothing
    Dim workbook As Excel.Workbook = Nothing
    Dim createdNewInstance As Boolean = False

    Try
      ' Try to get existing Excel instance first using COM API
      Try
        excelApp = CType(GetActiveObject("Excel.Application"), Excel.Application)
      Catch ex As Exception
        ' No existing instance found, create new one
        excelApp = New Excel.Application()
        excelApp.Visible = True
        createdNewInstance = True
      End Try

      ' Make sure Excel is visible
      If Not excelApp.Visible Then
        excelApp.Visible = True
      End If

      ' Check if workbook is already open in this Excel instance
      Dim isOpen As Boolean = False
      Try
        For Each wb As Excel.Workbook In excelApp.Workbooks
          If wb.FullName.Equals(excelFilePath, StringComparison.OrdinalIgnoreCase) Then
            workbook = wb
            isOpen = True
            Exit For
          End If
        Next
      Catch
        ' Workbooks collection may be empty
      End Try

      ' Open workbook if not already open
      If Not isOpen Then
        workbook = excelApp.Workbooks.Open(excelFilePath, [ReadOnly]:=True)
      End If

      ' Navigate to worksheet
      Dim worksheet As Excel.Worksheet = CType(workbook.Worksheets(worksheetName), Excel.Worksheet)
      worksheet.Activate()

      ' Select and scroll to cell
      Dim targetCell As Excel.Range = worksheet.Range(cellReference)
      targetCell.Select()

      ' Scroll to make cell visible
      Try
        excelApp.ActiveWindow.ScrollRow = Math.Max(1, targetCell.Row - 5)
        excelApp.ActiveWindow.ScrollColumn = Math.Max(1, targetCell.Column - 2)
      Catch
        ' ScrollRow/ScrollColumn might not be available in all contexts
      End Try

      ' Bring Excel window to front using Windows API
      Try
        Dim hwnd As IntPtr = New IntPtr(excelApp.Hwnd)
        SetForegroundWindow(hwnd)
      Catch
        ' If SetForegroundWindow fails, that's okay
      End Try

    Catch ex As Exception
      Throw New Exception("Error navigating to Excel cell: " & ex.Message, ex)
    Finally
      ' Clean up COM objects (but don't close Excel)
      If workbook IsNot Nothing Then
        Marshal.ReleaseComObject(workbook)
      End If
      If excelApp IsNot Nothing Then
        Marshal.ReleaseComObject(excelApp)
      End If
    End Try
  End Sub

  ''' <summary>
  ''' Gets a running COM object by ProgID (replacement for Marshal.GetActiveObject in .NET 8.0)
  ''' </summary>
  Private Shared Function GetActiveObject(progId As String) As Object
    Dim clsid As Guid = Guid.Empty
    Dim hr As Integer = CLSIDFromProgID(progId, clsid)

    If hr < 0 Then
      Marshal.ThrowExceptionForHR(hr)
    End If

    Dim unk As IntPtr = IntPtr.Zero
    hr = GetActiveObject(clsid, IntPtr.Zero, unk)

    If hr < 0 Then
      Marshal.ThrowExceptionForHR(hr)
    End If

    If unk = IntPtr.Zero Then
      Throw New COMException("No running instance found")
    End If

    Try
      Return Marshal.GetObjectForIUnknown(unk)
    Finally
      If unk <> IntPtr.Zero Then
        Marshal.Release(unk)
      End If
    End Try
  End Function

  ' Windows API declarations for COM interop
  <DllImport("ole32.dll")>
  Private Shared Function CLSIDFromProgID(<MarshalAs(UnmanagedType.LPWStr)> progId As String, ByRef clsid As Guid) As Integer
  End Function

  <DllImport("oleaut32.dll", PreserveSig:=True)>
  Private Shared Function GetActiveObject(ByRef rclsid As Guid, pReserved As IntPtr, ByRef ppunk As IntPtr) As Integer
  End Function

  ' Windows API to bring window to front
  <DllImport("user32.dll")>
  Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
  End Function
End Class