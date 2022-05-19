Attribute VB_Name = "Init"
Option Compare Database
Option Explicit

' References:
'   Microsoft Scripting Runtime

Dim fs As New FileSystemObject

' Creates property and personnel tables
Public Sub Init()

  Dim fs As New FileSystemObject
  Dim sql_file As TextStream
  Dim sql_str As Variant   ' SQL statements array
  
  Set sql_file = fs.OpenTextFile(CurrentProject.Path & "\tables.sql", ForReading, TristateFalse)
  sql_str = Split(sql_file.ReadAll(), "=sep=")
  
  ' tbl_personnel
  CurrentDb.Execute sql_str(0), dbFailOnError
  ' tbl_property
  CurrentDb.Execute sql_str(1), dbFailOnError

End Sub

' Imports CSV data into a respective table
Public Sub ImportData(table As String, csv_path As String)

  Dim csv_file As TextStream
  Dim csv_data As Variant
  Dim csv_row, csv_field As Long
  Dim new_record As DAO.Recordset

  Set csv_file = fs.OpenTextFile(csv_path, ForReading, TristateFalse)
  csv_data = ParseCSVToArray(csv_file.ReadAll())
  If IsNull(csv_data) Then
    Debug.Print "No CSV returned: " & Err.Number & " (" & Err.Source & ") " & Err.Description
  End If
  
  For csv_row = LBound(csv_data, 1) To UBound(csv_data, 1)
    Set new_record = CurrentDb.OpenRecordset(table)
    new_record.AddNew
    For csv_field = LBound(csv_data, 2) To UBound(csv_data, 2)
      new_record(csv_data(LBound(csv_data, 1), csv_field)) = csv_data(csv_row, csv_field)
    Next
    new_record.Update
  Next

End Sub

Sub test()
  ImportData "tbl_personnel", CurrentProject.Path & "/personnel.csv"
End Sub
