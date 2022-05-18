Attribute VB_Name = "Init"
Option Compare Database
Option Explicit

' References:
'   Microsoft Scripting Runtime

Public Sub Init()

  Dim fs As New FileSystemObject
  Dim sql_file As TextStream
  Dim sql_str As Variant   ' SQL statements array
  
  Set sql_file = fs.OpenTextFile(CurrentProject.Path & "\tables.sql", ForReading, TristateFalse)
  sql_str = Split(sql_file.readall(), "=sep=")
  
  ' tbl_personnel
  CurrentDb.Execute sql_str(0), dbFailOnError
  ' tbl_property
  CurrentDb.Execute sql_str(1), dbFailOnError

End Sub
