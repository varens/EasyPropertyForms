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
  Dim sql_str As Variant              ' SQL statements array
  Dim propfrm_query As DAO.QueryDef   ' Query object for the property-personnel form
  Dim propfrm_query_name As String    ' Name for the property-personnel query
  Dim namelist_query_name As String   ' Name for the query used in Assigned combobox
  Dim prop_pers_frm As Form           ' Property-personnel form object
  Dim prop_pers_frm_name As String
  Dim query_field As Field
  Dim ctrl_type As Integer
  
  propfrm_query_name = "qry_property_personnel"
  namelist_query_name = "qry_personnel_name_list"
  prop_pers_frm_name = "subfrm_property_personnel"
  
  ' Retrieve SQL statements
  Set sql_file = fs.OpenTextFile(CurrentProject.Path & "\tables.sql", ForReading, TristateFalse)
  sql_str = Split(sql_file.ReadAll(), "=sep=")
  
  ' tbl_personnel
  CurrentDb.Execute sql_str(0), dbFailOnError
  ' tbl_property
  CurrentDb.Execute sql_str(1), dbFailOnError
  ' tbl_property-tbl_personnel query
  Set propfrm_query = CurrentDb.CreateQueryDef(propfrm_query_name, sql_str(2))
  ' Personnel names query
  CurrentDb.CreateQueryDef namelist_query_name, sql_str(3)
  
  ' Create the property-personnel (sub)form
  Set prop_pers_frm = CreateForm()
  With prop_pers_frm
    .RecordSource = propfrm_query_name
    .DefaultView = 2                      ' Datasheet
    .AllowDatasheetView = True
    .AllowFormView = False
    .AllowLayoutView = False
    .Caption = "Property-Personnel"
  End With
  DoCmd.Restore
  DoCmd.Save acForm, prop_pers_frm.Name
  'DoCmd.Rename sub_form_name, acForm, sub_form.Name
  
  ' Iterate over property-personnel query object fields creating corresponding controls
  For Each query_field In propfrm_query.fields
    Select Case query_field.Name
      Case "personnel_id", "property_id": GoTo next_field
      Case "toggle":          ctrl_type = acCheckBox
      Case "assigned":        ctrl_type = acComboBox
      Case Else:              ctrl_type = acTextBox
    End Select
    CreateControl prop_pers_frm.Name, ctrl_type, acDetail, , query_field.Name
next_field:
  Next query_field
  Set query_field = Nothing

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
  
  For csv_row = LBound(csv_data, 1) + 1 To UBound(csv_data, 1)
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
  ImportData "tbl_property", CurrentProject.Path & "/property.csv"
End Sub
