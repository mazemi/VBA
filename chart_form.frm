VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} chart_form 
   Caption         =   "Chart Wizard"
   ClientHeight    =   4368
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   7800
   OleObjectBlob   =   "chart_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "chart_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub CommandCancel_Click()
    End
End Sub

Private Sub CommandNext_Click()
    If Me.OptionAll Then
        CHOSEN_CHART = 1
        If overal_chart_check Then
            Unload Me
            DISAGGREGATION_LEVEL = "ALL"
            DISAGGREGATION_VALUE = vbNullString
            DISAGGREGATION_LABEL = vbNullString
            Call generate_data_chart
        End If
        
    ElseIf Me.OptionDisaggregation Then
        CHOSEN_CHART = 2
        If other_chart_check Then
            Unload Me
            other_chart_form.Show
        End If
        
    ElseIf Me.OptionSingleChart Then
        If Not worksheet_exists("dm_backend") Or Not worksheet_exists("indi_list") Then
            MsgBox "Please first analyze the data, then try to generate charts. ", vbInformation
            Exit Sub
        End If
        CHOSEN_CHART = 3
        Unload Me
        single_chart_form.Show
    End If
End Sub


Private Sub OptionAll_Click()
    Call update_caption
End Sub

Private Sub OptionDisaggregation_Click()
    Call update_caption
End Sub

Private Sub OptionSingleChart_Click()
    Call update_caption
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .left = Application.left + (0.5 * Application.Width) - (0.5 * .Width)
        .top = Application.top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
    If CHOSEN_CHART = 1 Then
        Me.OptionDisaggregation.value = True
        Me.CommandNext.Caption = "Generate Chart"
    ElseIf CHOSEN_CHART = 2 Then
        Me.OptionDisaggregation.value = True
    ElseIf CHOSEN_CHART = 3 Then
        Me.OptionSingleChart.value = True
    End If
End Sub

Function overal_chart_check() As Boolean
    On Error GoTo ErrorHandler
    If Not worksheet_exists("result") Then
        MsgBox "Please first analyze the data, then try to generate charts", vbInformation
        overal_chart_check = False
        Exit Function
    End If
    
    If Not worksheet_exists("indi_list") Then
        MsgBox "Please first analyze the data with 'ALL' disaggrigation level." & vbCrLf & _
            "Then try to generate charts for overall data.", vbInformation
        overal_chart_check = False
        Exit Function
    End If
    
    Set dis_rng = sheets("indi_list").Range("G1").CurrentRegion
    
    For Each c In dis_rng
        If c = "ALL" Then
            has_all_dis = True
        End If
    Next c
    
    If Not has_all_dis Then
        MsgBox "Please first analyze the data with 'ALL' disaggrigation level." & vbCrLf & _
            "Then try to generate charts for overall data.", vbInformation
        overal_chart_check = False
        Exit Function
    End If
    
    overal_chart_check = True
    Exit Function
    
ErrorHandler:

    MsgBox "Please first analyze the data with 'ALL' disaggrigation level." & vbCrLf & _
        "Then try to generate charts for overall data NOTE2.", vbInformation
    overal_chart_check = False
    
End Function

Function other_chart_check() As Boolean
    On Error GoTo ErrorHandler
    If Not worksheet_exists("result") Then
        MsgBox "Please first analyze the data, then try to generate charts", vbInformation
        other_chart_check = False
        Exit Function
    End If
    
    If Not worksheet_exists("indi_list") Then
        MsgBox "Please first analyze the data with desired disaggrigation level." & vbCrLf & _
            "Then try to generate charts.", vbInformation
        other_chart_check = False
        Exit Function
    End If
    
    Set dis_rng = sheets("indi_list").Range("G1").CurrentRegion
    
    For Each c In dis_rng
        If c <> "ALL" Then
            has_other_dis = True
        End If
    Next c
    
    If Not has_other_dis Then
        MsgBox "Please first analyze the data with desired disaggrigations level." & vbCrLf & _
            "Then try to generate charts.", vbInformation
        other_chart_check = False
        Exit Function
    End If
    
    other_chart_check = True
    Exit Function
    
ErrorHandler:

    MsgBox "Please first analyze the data with desired disaggrigations level." & vbCrLf & _
        "Then try to generate charts. NOTE2.", vbInformation
    other_chart_check = False
    
End Function

Sub update_caption()
    Dim button_caption As String
    If Me.OptionAll.value Then
        button_caption = "Generate Chart"
    ElseIf Me.OptionDisaggregation.value Then
        button_caption = "Next"
    ElseIf Me.OptionSingleChart.value Then
        button_caption = "Next"
    End If
    
    Me.CommandNext.Caption = button_caption
End Sub
