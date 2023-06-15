VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} weighting_frm 
   Caption         =   "Weighting Setting"
   ClientHeight    =   5592
   ClientLeft      =   84
   ClientTop       =   390
   ClientWidth     =   8826.001
   OleObjectBlob   =   "weighting_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "weighting_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim sheet_li As Collection
    Set sheet_li = sheet_list 'Get the collection of worksheet names
    Dim sh As Variant 'name of a sheet
    For Each sh In sheet_li
        If ThisWorkbook.Worksheets(CStr(sh)).Visible Then
            Me.CombData.AddItem sh
            Me.ComboSampling.AddItem sh
        End If
    Next sh
    
    Me.ComboPopulation.Enabled = False
    Me.ComboSamplingStrata.Enabled = False
    Me.ComboDataStrata.Enabled = False
End Sub

Function sheet_list() As Collection
    'This function returns a collection of worksheet names in the workbook
    Dim sheets As Collection
    Set sheets = New Collection
    Dim ws As Worksheet 'Use a worksheet variable instead of an index
    For Each ws In ThisWorkbook.Worksheets 'Loop through each worksheet in the workbook
        sheets.Add ws.Name 'Add the worksheet name to the collection
    Next ws
    Set sheet_list = sheets
End Function

Private Sub PopulateComboBox(sheet_name As String, con As String)
'    On Error Resume Next
    Dim header_arr() As Variant
    Dim c As Control
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(sheet_name)
    
    ' array of header (list of questions)
    header_arr = ws.Range(ws.Cells(1, 1), ws.Cells(1, 1).End(xlToRight)).Value2
    
    For Each c In Me.Controls
        If c.Name = con Then
            c.Clear
            For Each i In header_arr

                c.AddItem i
            Next
        End If
    Next
End Sub

Private Sub ComboSampling_Change()
    'This subroutine updates the population and sampling strata combo boxes based on the selected worksheet name
    Dim val As String
    val = Me.ComboSampling.Value
    
    Me.ComboPopulation.Enabled = True
    Me.ComboPopulation.Clear
    
    Me.ComboSamplingStrata.Enabled = True
    Me.ComboSamplingStrata.Clear
    
    Call PopulateComboBox(val, "ComboPopulation")
    Call PopulateComboBox(val, "ComboSamplingStrata")
End Sub

Private Sub CombData_Change()
    Dim val As String
    val = Me.CombData.Value

    Me.ComboDataStrata.Enabled = True
    Me.ComboDataStrata.Clear
    
    Call PopulateComboBox(val, "ComboDataStrata")
End Sub



