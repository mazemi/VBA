Attribute VB_Name = "Tab_module"

'Callback for KOBOSetting onAction
Sub KOBOSetting(control As IRibbonControl)
    On Error Resume Next
    setting_form.Show
End Sub

'Callback for CancelDo onAction
Sub CancelDo(control As IRibbonControl)
    On Error Resume Next
    Application.DisplayAlerts = False
    
    Application.ScreenUpdating = False
    Application.StatusBar = False
    
    If worksheet_exists("keen") Then
        sheets("keen").Visible = xlSheetHidden
        sheets("keen").Delete
    End If
    
    If worksheet_exists("keen2") Then
        sheets("keen2").Visible = xlSheetHidden
        sheets("keen2").Delete
    End If
    
    If worksheet_exists("indi_list") Then
        sheets("indi_list").Visible = xlSheetHidden
        sheets("indi_list").Delete
    End If
    
    If worksheet_exists("temp_sheet") Then
        sheets("temp_sheet").Visible = xlSheetHidden
        sheets("temp_sheet").Delete
    End If
    
    If worksheet_exists("redeem") Then
        sheets("redeem").Visible = xlSheetHidden
        sheets("redeem").Delete
    End If
    
    End
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
     
End Sub

'Callback for DownloadAudit onAction
Sub DownloadAudit(control As IRibbonControl)
    On Error Resume Next
    Call download_audit
End Sub

'Callback for TimeChecking onAction
Sub TimeChecking(control As IRibbonControl)
    Call time_check
End Sub

'Callback for EmptyColumns onAction
Sub EmptyColumns(control As IRibbonControl)
    Call no_value_col
End Sub

'Callback for CheckDuplicates onAction
Sub CheckDuplicates(control As IRibbonControl)
    On Error Resume Next
    Call find_duplicate
End Sub

'Callback for AddLabel onAction
Sub addLabel(control As IRibbonControl)
    On Error Resume Next
    Call add_label
End Sub

'Callback for ClearFilter onAction
Sub ClearFilter(control As IRibbonControl)
    On Error Resume Next
    Call clear_active_filter
End Sub

'Callback for AutoCheckData onAction
Sub AutoCheckData(control As IRibbonControl)
    On Error Resume Next
    Call auto_check
End Sub

'Callback for SetLogicalChekc onAction
Sub SetLogicalChekc(control As IRibbonControl)
    On Error Resume Next
    plan_form.Show
End Sub

'Callback for LogicalChekcList onAction
Sub LogicalChekcList(control As IRibbonControl)
    On Error Resume Next
    plan_list_form.Show
End Sub

'Callback for ImportLogicalChekc onAction
Sub ImportLogicalChekc(control As IRibbonControl)
    On Error Resume Next
    Call import_plan
End Sub

'Callback for ExportLogicalChekc onAction
Sub ExportLogicalChekc(control As IRibbonControl)
    On Error Resume Next
    Call export_plan
End Sub


'Callback for ConsistencyCheck onAction
Sub ConsistencyCheck(control As IRibbonControl)
    On Error Resume Next
    Call consistency_check
End Sub

'Callback for AddToLogs onAction
Sub AddToLogs(control As IRibbonControl)
    On Error Resume Next
    Call pattern_check(False)
End Sub


'Callback for AddToLogsMore onAction
Sub AddToLogsMore(control As IRibbonControl)
    On Error Resume Next
    extra_logs_form.Show
End Sub

'Callback for CheckLogDuplicate onAction
Sub CheckLogDuplicate(control As IRibbonControl)
    On Error Resume Next
    Call find_duplicate_log
End Sub

'Callback for DetectOutliers onAction
Sub DetectOutliers(control As IRibbonControl)
    On Error Resume Next
    Call calulate_quartiles
End Sub

'Callback for ReplaceLogs onAction
Sub ReplaceLogs(control As IRibbonControl)
    On Error Resume Next
    Call replace_log
End Sub

'Callback for Weighting onAction
Sub DoWeighting(control As IRibbonControl)
    On Error Resume Next
    weighting_form.Show
End Sub

'Callback for Disaggregations onAction
Sub Disaggregations(control As IRibbonControl)
    On Error Resume Next
    disaggregation_form.Show
End Sub

'Callback for RunAnalysis onAction
Sub RunAnalysis(control As IRibbonControl)
    On Error Resume Next
    analysis_form.Show
End Sub

'Callback for AllFigures onAction
Sub Figures(control As IRibbonControl)
    On Error Resume Next
    chart_form.Show
'    Call generate_data_chart
End Sub

'Callback for SingleChart onAction
Sub SingleChart(control As IRibbonControl)
    On Error Resume Next
    single_chart_form.Show
End Sub

'Callback for FindIndicator onAction
Sub FindIndicator(control As IRibbonControl)
    On Error Resume Next
    find_form.Show
End Sub


