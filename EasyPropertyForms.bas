Attribute VB_Name = "EasyPropertyForms"
Option Compare Database
Option Explicit

' Use cases:
'   - print DA 3749 for all personnel
'   - print DA 3749 individually

' This modules relies on a sanitized blank PDF of DA Form 3749
' One way to sanitize the official PDF is with pdftk
' The expected location is LOCATION_OF_ACCESSDB\assets\DA3749.pdf
' The module saves all output to \output\DA3749_live.pdf
'
' One challenge of working with PDF forms is maintaining uniqueness of field names
' Simply duplicating pages with forms is not practical as most PDF client software
' will complain or will not reliably handle such documents as each field must have
' a unique ID or name. The limitation is overcome here by flattening each page as
' it gets populated before copying/inserting a blank page from the original file.
' Coordinates of signature fields are saved beforehand and used to create new ones
' on a "flat" page.

Dim DA3749_PATH As String
Dim DA3749_WORKING_PATH As String
Dim OUT_DIR As String
Dim field_name_keys(7) As String
Dim field_name_suffix As Variant
Dim acro_app As Acrobat.CAcroApp
Dim pdf_3749 As Acrobat.CAcroPDDoc
Dim grade_for


Public Sub init()

  Dim record
  Dim cbrn_recordset As DAO.recordset
  Dim fs As New FileSystemObject
  Dim total_records As Long
  Dim record_count As Integer

  On Error GoTo Croak

  DA3749_PATH = CurrentProject.Path & "\assets\DA3749.pdf"
  OUT_DIR = CurrentProject.Path & "\output"
  DA3749_WORKING_PATH = OUT_DIR & "\DA3749_live.pdf"
  Set acro_app = CreateObject("AcroExch.App")
  Set pdf_3749 = CreateObject("AcroExch.PDDoc")

  populate_globals

  If Not fs.FileExists(DA3749_WORKING_PATH) Then _
    fs.CopyFile DA3749_PATH, DA3749_WORKING_PATH

  Set record = CreateObject("Scripting.Dictionary")
  Set cbrn_recordset = CurrentDb.OpenRecordset("qry_m50", , dbReadOnly)

  If cbrn_recordset.BOF And cbrn_recordset.EOF Then
    MsgBox "Database returned no records"
    Exit Sub
  End If

  With cbrn_recordset
    .MoveLast
    total_records = .RecordCount
    .MoveFirst

    record_count = 1
    Do Until .EOF
      record.Add "RECEIPT", Split(![admin], "-")(2)
      record.Add "STOCK", ![nsn]
      record.Add "SERIAL", ![serial]
      record.Add "DESCRIPT", "Mask, Chemical-Biological, M50 (Size: " & ![mask_size] & ")"
      record.Add "NAME", ![person_name]
      If ![designation] = "MIL" Then
        record.Add "GRADE", grade_for(CStr(![rank]))
      Else
        record.Add "GRADE", ![designation]
      End If
      record.Add "UNIT", "Some unit, " & ![Section]
      record.Add "FROM", "the CBRN room"

      write_3749 record, record_count = total_records

      record.RemoveAll
      record_count = record_count + 1
      .MoveNext
    Loop
  End With

  Set cbrn_recordset = Nothing
  Set record = Nothing
  Set pdf_3749 = Nothing
  Set acro_app = Nothing

  MsgBox "Generated " & total_records & " DA 3749 forms."

  Exit Sub

Croak:
  MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description

End Sub


Private Sub write_3749(record, is_last As Boolean)

  Dim pdf_js As Object
  Dim empty_form As Integer         ' indicates first empty 3749 on a page
  Dim i As Integer

  pdf_3749.Open (DA3749_WORKING_PATH)
  Set pdf_js = pdf_3749.GetJSObject
  empty_form = get_empty_index(pdf_js)

  If empty_form < 0 Then
    If empty_form = -1 Then flatten_form pdf_js
    pdf_js.InsertPages -1, DA3749_PATH
    empty_form = 0
  End If

  ' populate the form
  For i = 0 To UBound(field_name_keys)
    pdf_js.getField(get_field_name(field_name_keys(i), empty_form)).Value = _
        CStr(record.Item(field_name_keys(i)))
    ' set a monospaced typeface for numeric values
    If field_name_keys(i) = "RECEIPT" _
        Or field_name_keys(i) = "STOCK" _
        Or field_name_keys(i) = "SERIAL" Then
      pdf_js.getField(get_field_name(field_name_keys(i), empty_form)).TextFont = _
          "Consolas"
    End If
  Next i

  If is_last Then _
    flatten_form pdf_js

  pdf_3749.Save PDSaveFull, DA3749_WORKING_PATH

  Set pdf_js = Nothing
  pdf_3749.Close

  Exit Sub

End Sub


Private Function get_field_name(ByVal field_name_key As String, form_index As Integer)

' DA 3749 fields:
'   form1[0].Page1[0].UNIT[0]                     .UNIT(_[BCD])?
'   form1[0].Page1[0].RECEIPT[0]                  Admin number
'   form1[0].Page1[0].STOCK[0]                    NSN
'   form1[0].Page1[0].SERIAL[0]
'   form1[0].Page1[0].DESCRIPT[0]                 DESCRPT for form_index > 0
'   form1[0].Page1[0].FROM[0]
'   form1[0].Page1[0].NAME[0]
'   form1[0].Page1[0].signature_BUTTON1[0]        _BUTTON[1357]
'   form1[0].Page1[0].GRADE[0]
'   form1[0].Page1[0].signature_BUTTON2[0]        _BUTTON[2468]

  ' several field names are not consistent: DESCRIPT/DESCRPT
  If form_index > 0 And field_name_key = "DESCRIPT" Then _
    field_name_key = "DESCRPT"

  get_field_name = "form1[0].Page1[0]." & field_name_key & _
      field_name_suffix(form_index)

End Function

Private Sub populate_globals()

  Dim keys_tmp As Variant
  Dim i As Integer
  Dim ranks, grades

  keys_tmp = Array("UNIT", "RECEIPT", "STOCK", "SERIAL", "DESCRIPT", _
      "FROM", "NAME", "GRADE")
  For i = 0 To 7
    field_name_keys(i) = keys_tmp(i)
  Next i

  field_name_suffix = Array("[0]", "_B[0]", "_C[0]", "_D[0]")

  ranks = Array("PVT", "PV2", "PFC", "SPC", "CPL", "SGT", "SSG", "SFC", "MSG", _
      "1SG", "SGM", "CSM", "2LT", "1LT", "CPT", "MAJ", "LTC", "COL", "CW1", _
      "CW2", "CW3", "CW4", "CW5")
  grades = Array("E-1", "E-2", "E-3", "E-4", "E-4", "E-5", "E-6", "E-7", "E-8", _
      "E-8", "E-9", "E-9", "O-1", "O-2", "O-3", "O-4", "O-5", "O-6", "W-1", _
      "W-2", "W-3", "W-4", "W-5")
  Set grade_for = CreateObject("Scripting.Dictionary")
  For i = 0 To UBound(ranks)
    grade_for.Add ranks(i), grades(i)
  Next i

End Sub

Private Sub flatten_form(pdf_js As Object)

  Dim sig_rects As New Collection
  Dim i, j As Integer
  Dim sig_field_name As String
  Dim sig_rect

  ' collect signature fields coordinates
  For i = 1 To 8
    sig_rects.Add pdf_js.getField("form1[0].Page1[0].signature_BUTTON" & i & "[0]").rect, "r" & i
  Next i

  pdf_js.flattenPages 0

  ' create new signature fields
  For i = 1 To 8
    sig_field_name = "proper.signature." & pdf_js.numPages & "." & i
    ' enlarge the rectangle
    sig_rect = sig_rects.Item("r" & i)
    sig_rect(3) = sig_rect(3) - 10
    pdf_js.addField sig_field_name, "signature", 0, Split(Join(sig_rect))
    sig_rects.Remove "r" & i
  Next i

  Set sig_rect = Nothing

End Sub

' A naive search for the first empty form on a page
' Admin num is assumed populated
' Return a 0-based index, -1 if entire page is filled, or -2 if no form found,
' likely because all pages are flat
Private Function get_empty_index(pdf_js As Object)

  Dim fields(3) As String
  Dim i As Integer

  fields(0) = "form1[0].Page1[0].RECEIPT[0]"
  fields(1) = "form1[0].Page1[0].RECEIPT_B[0]"
  fields(2) = "form1[0].Page1[0].RECEIPT_C[0]"
  fields(3) = "form1[0].Page1[0].RECEIPT_D[0]"

  get_empty_index = -1

  For i = 0 To 3
    ' field not found
    If VarType(pdf_js.getField(fields(i))) = 0 Then
      get_empty_index = -2
      GoTo Finish
    End If
    ' found an empty one
    If Len(pdf_js.getField(fields(i)).Value) = 0 Then
      get_empty_index = i
      GoTo Finish
    End If
  Next

Finish:

End Function
