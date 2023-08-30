VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} disaggregation_form 
   Caption         =   "Analysis Disaggregations"
   ClientHeight    =   4998
   ClientLeft      =   84
   ClientTop       =   390
   ClientWidth     =   9060.001
   OleObjectBlob   =   "disaggregation_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "disaggregation_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub set_indicator_validator()
        
    Dim dt_name As String
    dt_name = find_main_data
    
    ActiveWorkbook.sheets("analysis_list").Activate
    ActiveWorkbook.sheets("analysis_list").Range("A2:A" & rows.count).Select
    
    With sheets("analysis_list").Range("A2:A" & rows.count).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
             xlBetween, Formula1:="=" & dt_name & "!$1:$1"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = "Please enter a valid indicator."
        .ShowInput = True
        .ShowError = True
    End With
        
End Sub

' this function highlight the valid question type
Sub format_question_type()
    Dim conditional_rng As Range
    Dim condition1 As FormatCondition
    Dim condition2 As FormatCondition
    Dim condition3 As FormatCondition
    
    Set conditional_rng = sheets("analysis_list").Range("B2:B" & rows.count)
    
    'to clear existing conditional formatting
    conditional_rng.FormatConditions.Delete

    'to specify the condition for each format
    Set condition1 = conditional_rng.FormatConditions.Add(xlCellValue, xlEqual, "=""integer""")
    Set condition2 = conditional_rng.FormatConditions.Add(xlCellValue, xlEqual, "=""decimal""")
    Set condition3 = conditional_rng.FormatConditions.Add(xlCellValue, xlEqual, "=""select_one""")
    Set condition4 = conditional_rng.FormatConditions.Add(xlCellValue, xlEqual, "=""select_multiple""")
    
    With condition1
        .Font.Color = RGB(0, 176, 59)
    End With
    
    With condition2
        .Font.Color = RGB(0, 176, 59)
    End With
    
    With condition3
        .Font.Color = RGB(0, 176, 59)
    End With
    
    With condition4
        .Font.Color = RGB(0, 176, 59)
    End With

End Sub

Private Sub CommandReset_Click()
    sheets("dissagregation_setting").Cells.Clear
    sheets("dissagregation_setting").Cells(1, 1) = "Disaggregation Level"
    sheets("dissagregation_setting").Cells(1, 2) = "Weight"
    Call referesh_list
End Sub

Private Sub CommandSave_Click()
    Application.ScreenUpdating = False
    Dim val As String
    Dim current_dt_name As String
    Dim new_name As String
    
    If Me.ComboSheets.value <> vbNullString Then
        current_dt_name = Me.ComboSheets.value
        new_name = alpha_numeric_only(current_dt_name)
        
        If Len(new_name) > 15 Then
            new_name = left(new_name, 15)
        End If
        
        Public_module.DATA_SHEET = new_name
        dt_sheet = SaveRegistrySetting("ramSetting", "dataReg", new_name)
        sheets(current_dt_name).Name = new_name
        
    Else
        Public_module.DATA_SHEET = Me.ComboSheets.value
        dt_sheet = SaveRegistrySetting("ramSetting", "dataReg", Me.ComboSheets.value)
    End If
    

    Public_module.DATA_SHEET = new_name
    dt_sheet = SaveRegistrySetting("ramSetting", "dataReg", new_name)
    
    If Me.ListWeight.ListCount > 0 Then
        ' check if dissagregation_setting sheet exist
        If Not worksheet_exists("analysis_list") Then
            Call create_sheet("dissagregation_setting", "analysis_list")
            sheets("analysis_list").Cells(1, 1) = "question"
            sheets("analysis_list").Cells(1, 2) = "type"
            
            sheets("analysis_list").columns("A:A").ColumnWidth = 30
            sheets("analysis_list").columns("B:B").ColumnWidth = 20
        
            With sheets("analysis_list").Range("A1:B1").Interior
                .Pattern = xlSolid
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.15
                .Parent.Font.Bold = True
            End With
            
            Call set_indicator_validator
            Call format_question_type
            
            ' Call add_note
                   
        End If
        
        If Me.CheckBoxAll Then
            Call add_all_indicators
            With sheets("analysis_list")
                .columns("B:B").Copy
                .columns("B:B").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, transpose:=False
            End With
            sheets("analysis_list").Activate
        End If
        
    End If
      
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Unload Me
End Sub

Sub add_all_indicators()

    Dim sur_ws As Worksheet
    Dim tmp_ws As Worksheet
    Dim analys_ws As Worksheet
    Dim rng As Range
    If worksheet_exists("temp_sheet") <> True Then
        Call create_sheet(find_main_data, "temp_sheet")
    End If
    
    Set sur_ws = ThisWorkbook.sheets("xsurvey")
    Set tmp_ws = sheets("temp_sheet")
    Set analys_ws = sheets("analysis_list")
    tmp_ws.Cells.ClearContents
    Set rng = sur_ws.Range("A1").CurrentRegion
    
    tmp_ws.Cells(1, 1) = "type"
    tmp_ws.Cells(2, 1) = "integer"
    tmp_ws.Cells(3, 1) = "decimal"
    tmp_ws.Cells(4, 1) = "select_one *"
    tmp_ws.Cells(5, 1) = "select_multiple *"
    tmp_ws.Cells(6, 1) = "calculate"
    tmp_ws.Cells(1, 3) = "name"
    
    rng.AdvancedFilter xlFilterCopy, tmp_ws.Range("A1").CurrentRegion, tmp_ws.Range("C1")
    analys_ws.Range(analys_ws.Range("A2:B2"), analys_ws.Range("A2:B2").End(xlDown)).ClearContents
    last_indicator = tmp_ws.Cells(rows.count, 3).End(xlUp).row
    analys_ws.Range("A2:A" & last_indicator).Value2 = tmp_ws.Range("C2:C" & last_indicator).Value2
    analys_ws.Range("B2:B" & last_indicator).Formula = "=question_type(A2)"
    
    Application.DisplayAlerts = False
    sheets("temp_sheet").Delete
    Application.DisplayAlerts = True

End Sub


Private Sub UserForm_Initialize()
    On Error GoTo err_handler
    Application.ScreenUpdating = False
    
    ' check if dissagregation_setting sheet exist
    If Not worksheet_exists("dissagregation_setting") Then
        Call create_sheet(sheets(sheets.count).Name, "dissagregation_setting")
        sheets("dissagregation_setting").Cells(1, 1) = "Disaggregation Level"
        sheets("dissagregation_setting").Cells(1, 2) = "Weight"
        Worksheets("dissagregation_setting").visible = xlVeryHidden
    End If
    Dim dt_sheet_name As String

    
    ' check if dissagregation_setting sheet exist
    If ThisWorkbook.Worksheets("xsurvey").Range("A1") = vbNullString Then
        MsgBox "Please import the KOBO tools.    ", vbInformation
        End
    End If
    
    dt_sheet_name = find_main_data
  
    If worksheet_exists(dt_sheet_name) Then
        Me.ComboSheets.Text = dt_sheet_name
    End If

    Dim dissagregation_ws As Worksheet, dis_rng As Range
    Dim sheet_li As Collection
    Set sheet_li = sheet_list                    'Get the collection of worksheet names
    Dim sh As Variant                            'name of a sheet
    For Each sh In sheet_li
        If ActiveWorkbook.Worksheets(CStr(sh)).visible Then
            If CStr(sh) <> "result" And CStr(sh) <> "log_book" And CStr(sh) <> "analysis_list" And _
               CStr(sh) <> "dissagregation_setting" And CStr(sh) <> "overall" And CStr(sh) <> "survey" And _
               CStr(sh) <> "keen" And CStr(sh) <> "indi_list" And CStr(sh) <> "temp_sheet" And _
               CStr(sh) <> "choices" And CStr(sh) <> "xsurvey_choices" And CStr(sh) <> "datamerge" Then
                Me.ComboSheets.AddItem sh
            End If
        End If
    Next sh
    
    '    Me.ComboQuestions.Enabled = False
    '    Me.ComboWeight.Enabled = False
    
    Set dissagregation_ws = sheets("dissagregation_setting")

    Set dis_rng = dissagregation_ws.Range("A1:B" & dissagregation_ws.Range("A" & dissagregation_ws.rows.count).End(xlUp).row)
    
    With Me.ListWeight
        .BorderStyle = 1
        .ColumnHeads = True
        .columnCount = dis_rng.columns.count
        .ColumnWidths = "120,20"
        .RowSource = dis_rng.Parent.Name & "!" & dis_rng.Resize(dis_rng.rows.count - 1).Offset(1).Address
    End With
     
    Me.LabelTool.Caption = "Integrated Tool: " & vbCrLf & GetRegistrySetting("ramSetting", "koboToolReg")
     
    Me.ComboQuestions.Enabled = True
    
    Me.ComboWeight.Enabled = True
    
    Application.ScreenUpdating = True
    Exit Sub
     
err_handler:

    If Not worksheet_exists("dissagregation_setting") Then
        Call create_sheet(main_ws.Name, "dissagregation_setting")
        sheets("dissagregation_setting").Cells(1, 1) = "Disaggregation Level"
        sheets("dissagregation_setting").Cells(1, 2) = "Weight"
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub CommandAddWeight_Click()
    
    q = Me.ComboQuestions.value
    w = Me.ComboWeight.value
    
    If q = "" Or w = "" Then
        Exit Sub
    End If
    
    Dim rng As Range
    last_dis = sheets("dissagregation_setting").Cells(rows.count, 1).End(xlUp).row
    Set rng = sheets("dissagregation_setting").Range("A2:B" & CStr(last_dis))
    
    If rng.row > 1 Then
        For Each diss_value In rng.columns(1).Cells
            If diss_value = q Then
                MsgBox "Duplicate disaggregation!              ", vbExclamation
                Exit Sub
            End If
        Next
    End If
    
    If rng.row > 10 Then
        MsgBox "Maximum disaggregation levels!              ", vbExclamation
        Exit Sub
    End If
    
    sheets("dissagregation_setting").Cells(last_dis + 1, 1) = Me.ComboQuestions.value
    sheets("dissagregation_setting").Cells(last_dis + 1, 2) = Me.ComboWeight.value
    
    Call referesh_list
End Sub

Private Sub ComboSheets_Change()
    Dim val As String
    
    val = Me.ComboSheets.value
    
    If val = "" Then
        Exit Sub
    End If

    Me.ComboQuestions.Enabled = True
    Me.ComboQuestions.Clear
    
    Me.ComboWeight.Enabled = True
    Me.ComboWeight.Clear
    
    Me.ComboWeight.AddItem ("yes")
    Me.ComboWeight.AddItem ("no")
    
    Call PopulateComboBox(val, "ComboQuestions")

End Sub

Private Sub PopulateComboBox(sheet_name As String, con As String)
    '    On Error Resume Next
    Dim header_arr() As Variant
    Dim c As control
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.sheets(sheet_name)
       
    header_arr = ws.Range(ws.Cells(1, 1), ws.Cells(1, 1).End(xlToRight)).Value2
    
    Dim not_for_dis As New Collection
    
    '    Set not_for_dis = not_good_dis

    For Each c In Me.Controls
        If c.Name = con Then
            c.Clear
            c.AddItem "ALL"
            For Each i In header_arr
                c.AddItem i
            Next
        End If
    Next
End Sub

Sub referesh_list()
    On Error Resume Next
    Dim dissagregation_ws As Worksheet
    Dim dis_rng As Range
    Set dissagregation_ws = sheets("dissagregation_setting")

    Set dis_rng = dissagregation_ws.Range("A1:B" & dissagregation_ws.Range("A" & dissagregation_ws.rows.count).End(xlUp).row)
    
    With Me.ListWeight
        .BorderStyle = 1
        .ColumnHeads = True
        .columnCount = dis_rng.columns.count
        .ColumnWidths = "120,20"
        .RowSource = dis_rng.Parent.Name & "!" & dis_rng.Resize(dis_rng.rows.count - 1).Offset(1).Address
    End With
End Sub

'Private Sub CommandRemove_Click()
'
''    On Error GoTo err_handler
'    Dim strng As String
'    Dim lCol As Long, lRow As Long
'
'    With Me.ListWeight
'        For lRow = 0 To .ListCount - 1           '<--| loop through listbox rows
'            If .Selected(lRow) Then              '<--| if current row selected
'                For lCol = 0 To .columnCount - 1 '<--| loop through listbox columns
'                    strng = strng & .List(lRow, lCol) & "," '<--| build your output string
'                Next lCol
'                Exit For
'            End If
'        Next lRow
'    End With
'
'    Dim Result() As String
'    Result() = Split(strng, ",")
'
'    Dim a As Integer
'    a = 1
'
'    Do While Cells(a, 1) <> ""
'
'        If sheets("dissagregation_setting").Cells(a, 1) = Result(0) And sheets("dissagregation_setting").Cells(a, 2) = Result(1) Then
'            sheets("dissagregation_setting").rows(a).Delete Shift:=xlUp
'            'Row counter should not be incremented if row was just deleted
'        Else
'            'Increment a for next row only if row not deleted
'            a = a + 1
'        End If
'
'    Loop
'
'    Call referesh_list
''    Exit Sub
''
''err_handler:
'
'End Sub

Function not_good_dis() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets("xsurvey")
    Dim rng As Range
    Dim not_good_collection As New Collection
    last_row_survey = ws.Cells(ws.rows.count, 1).End(xlUp).row
    
    ws.Range("$A$2:$B$" & CStr(last_row_survey)).AutoFilter Field:=1, Criteria1:="<>*select_one*", Operator:=xlAnd, Criteria2:="<>*calculate*"
    Set rng = ws.Range("B2:B" & last_row_survey).SpecialCells(xlCellTypeVisible)
        
    For Each d In rng
        not_good_collection.Add d
    Next
    
    Set not_good_dis = not_good_collection
End Function

' add a note for a quick quide
Sub add_note()

    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 420, 100, 500, 90).Select
        
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset2
    Selection.ShapeRange.Adjustments.item(1) = 0.02057
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = _
                                                                 "Please add the necessary indicator in the first column.You can use the question_type formula in the second column, for example: =question_type(A2)" & _
                                                                 " Valid types are integer, decimal, select_one and select_multiple."
        
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 148). _
         ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignLeft
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 73).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
        .Size = 11
        .Name = "+mn-lt"
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(74, 14).Font
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
        .Size = 11
        .Name = "+mn-lt"
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(88, 43).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
        .Size = 11
        .Name = "+mn-lt"
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(131, 18).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
        .Size = 11
        .Name = "+mn-lt"
    End With

    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 1")).Select
    Selection.ShapeRange.IncrementLeft -25.5
    Selection.ShapeRange.IncrementTop -87
    
    ActiveSheet.Cells(2, 1).Select
End Sub


