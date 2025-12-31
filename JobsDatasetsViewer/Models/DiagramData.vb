' JSON Data Model Classes
' These match the JSON schema from the generator application

Public Class DiagramMetadata
  Public Property projectName As String
  Public Property generatedDate As String
  Public Property excelFilePath As String
  Public Property selectedDatasetTypes As List(Of String)
  Public Property applicationVersion As String
End Class

Public Class ExcelReference
  Public Property worksheet As String
  Public Property row As Integer
  Public Property column As String
  Public Property cellReference As String
End Class

Public Class VisualProperties
  Public Property color As String
  Public Property objectType As String
End Class

Public Class DatasetInfo
  Public Property name As String
  Public Property type As String
  Public Property relationship As String
  Public Property excelReference As ExcelReference
  Public Property visualProperties As VisualProperties
End Class

Public Class JobInfo
  Public Property name As String
  Public Property datasets As List(Of DatasetInfo)
End Class

Public Class DiagramData
  Public Property metadata As DiagramMetadata
  Public Property jobs As List(Of JobInfo)
End Class