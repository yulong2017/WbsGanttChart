Option Explicit

Const g_max_level As Integer = 5

Dim g_level_number(g_max_level) As Integer

Dim g_task_startday_min(g_max_level) As Date
Dim g_task_endday_max(g_max_level) As Date
Const g_task_default_date As Date = #12/31/9999#

Dim g_task_actday(g_max_level) As Single
Dim g_task_planday(g_max_level) As Single

Const g_level1_format As Integer = 0
Const g_level2_format As Integer = 4
Const g_level3_format As Integer = 8
Const g_level4_format As Integer = 12
Const g_level5_format As Integer = 16

Const g_start_day_x As Integer = 3
Const g_start_day_y As Integer = 26
Const g_ganttchart_start_col As Integer = 26

Const g_task_area_start_line As Integer = 6
Const g_task_area_start_col As Integer = 2
Dim g_task_area_end_line As Integer
Const g_task_area_end_col As Integer = 25

Const g_task_name_col As Integer = 3
Const g_task_days_col As Integer = 6
Const g_task_start_day_col As Integer = 7
Const g_task_end_day_col As Integer = 8
Const g_task_process_col As Integer = 14
Const g_PrivateTask_col As Integer = 15
Const g_task_baseline_start_day_col As Integer = 16
Const g_task_baseline_end_day_col As Integer = 17
Const g_task_type_col As Integer = 24
Const g_task_level_col As Integer = 25

Dim g_date_end_col As Integer

Dim g_current_selection_row As Integer
Dim g_current_selection_col As Integer

Const g_GanttChartParentTaskColor As Single = -65536

Sub SaveCurrentSelectionPos()
    g_current_selection_row = Selection.Row
    g_current_selection_col = Selection.Column
End Sub

Sub UpdateCurrentSelectionPos()
    Range(Cells(g_current_selection_row, g_current_selection_col), Cells(g_current_selection_row, g_current_selection_col)).Select
End Sub

Sub test()
    MsgBox Cells(24, 14).Text
End Sub

Function GetGanttMinDate() As Date
    GetGanttMinDate = Cells(g_start_day_x, g_start_day_y).Value
End Function

Function GetGanttMaxDate() As Date
    GetGanttMaxDate = DateAdd("m", 6, Cells(g_start_day_x, g_start_day_y).Value)
End Function

Sub getTaskArea()
    Dim line As Integer
    Dim col As Integer

    line = g_task_area_start_line
    Do While Cells(line, g_task_name_col).Value <> ""
        line = line + 1
    Loop
    g_task_area_end_line = line - 1
    
    g_date_end_col = g_ganttchart_start_col + GetGanttMaxDate() - GetGanttMinDate()
    
    For col = g_start_day_y + 1 To g_date_end_col
        Cells(g_start_day_x + 1, col).Value = Cells(g_start_day_x + 1, col - 1).Value + 1
    Next
    
End Sub

Sub UpdateGanttChart()
    Dim line As Integer
    Dim level As Integer
    
    Call SaveCurrentSelectionPos
    Application.ScreenUpdating = False
    
    Call getTaskArea
  
    Call clearLevelNumber
    Call clearGanttChart

    Range(Cells(g_task_area_start_line, g_task_area_start_col), Cells(g_task_area_end_line, g_task_area_end_col)).Select
    Selection.Font.Bold = False
    
    For line = g_task_area_start_line To g_task_area_end_line
        formatLevel (line)
    Next
    
    Call updateTaskInfo

    For line = g_task_area_start_line To g_task_area_end_line
        drawGanttChart (line)
    Next

    Call drawCurrentDayLine
    Call updatePrivateTasks

    Call setBaseLineDate

    Call UpdateCurrentSelectionPos
    Application.ScreenUpdating = True
    
End Sub

Sub formatLevel(line As Integer)
    Dim level As Integer
    
    level = checkLevel(Cells(line, g_task_name_col).Value)
    LevelNumberAdd (level)
    If InStr(LTrim(Cells(line, g_task_name_col).Value), " ") Then
        Cells(line, g_task_name_col).Value = createLevelHead(level) & createLevelNumber(level) & " " & Strings.Split(LTrim(Cells(line, g_task_name_col).Value), " ")(1)
    Else
        Cells(line, g_task_name_col).Value = createLevelHead(level) & createLevelNumber(level) & " " & LTrim(Cells(line, g_task_name_col).Value)
    End If
    
    Cells(line, g_task_level_col).Value = level
    Cells(line, g_task_type_col).Value = "T"
    If line > g_task_area_start_line Then
        If level > Cells(line - 1, g_task_level_col).Value Then
            Cells(line - 1, g_task_type_col).Value = "P"
            Range(Cells(line - 1, g_task_area_start_col), Cells(line - 1, g_task_area_end_col)).Select
            Selection.Font.Bold = True
        End If
    End If

End Sub

Function checkLevel(task_name As String) As Integer
    If Left(task_name, g_level5_format) = Space(g_level5_format) Then
       checkLevel = 5
    ElseIf Left(task_name, g_level4_format) = Space(g_level4_format) Then
       checkLevel = 4
    ElseIf Left(task_name, g_level3_format) = Space(g_level3_format) Then
       checkLevel = 3
    ElseIf Left(task_name, g_level2_format) = Space(g_level2_format) Then
       checkLevel = 2
    Else
       checkLevel = 1
    End If
End Function

Sub clearLevelNumber()
    Dim level  As Integer
    
    For level = 0 To g_max_level
        g_level_number(level) = 0
    Next
End Sub

Sub LevelNumberAdd(level As Integer)
    level = level - 1
    If level > g_max_level Then
        MsgBox "Level" & level & "Error"
    End If
    
    g_level_number(level) = g_level_number(level) + 1
    
    For level = level + 1 To g_max_level
        g_level_number(level) = 0
    Next
End Sub

Function createLevelNumber(level As Integer) As String
    Dim i As Integer
    level = level - 1
    If level > g_max_level Then
        MsgBox "Level" & level & "Error"
    End If
    
    createLevelNumber = createLevelNumber & g_level_number(i)
    For i = 1 To level
        createLevelNumber = createLevelNumber & "." & g_level_number(i)
    Next
End Function

Function createLevelHead(level As Integer) As String
    If level > g_max_level Then
        MsgBox "Level" & level & "Error"
    End If
    
    Select Case level
        Case 1
            createLevelHead = Space(0)
        Case 2
            createLevelHead = Space(g_level2_format)
        Case 3
            createLevelHead = Space(g_level3_format)
        Case 4
            createLevelHead = Space(g_level4_format)
        Case 5
            createLevelHead = Space(g_level5_format)
    End Select
End Function

Sub drawGanttChart(line As Integer)
    Dim offset As Integer
    Dim date_len As Integer
    Dim base As Variant
    Dim i As Integer
    
    If IsError(Range("G" & line)) = True Then
        Exit Sub
    End If
    
    If Not ((Cells(line, g_task_end_day_col).Value >= GetGanttMinDate() And Cells(line, g_task_end_day_col).Value <= GetGanttMaxDate()) _
        And (Cells(line, g_task_start_day_col).Value >= GetGanttMinDate() And Cells(line, g_task_start_day_col).Value <= GetGanttMaxDate())) Then
        Exit Sub
    End If
    
    
    base = Cells(g_start_day_x, g_start_day_y).Value
    date_len = Cells(line, g_task_end_day_col).Value - Cells(line, g_task_start_day_col).Value
    offset = Cells(line, g_task_start_day_col).Value - base
    
    If Cells(line, g_task_type_col).Value = "T" Then
        If Cells(line, g_task_process_col).Value = 1 Then
            Range("F2").Select
        ElseIf Cells(line, g_task_process_col).Value > 0 Then
            Range("J2").Select
        Else
            Range("H2").Select
        End If
        Selection.Copy
        
        For i = 0 To date_len
            Cells(line, g_ganttchart_start_col + offset + i).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                                      SkipBlanks:=False, Transpose:=False
        Next
        
        Range(Cells(line, g_ganttchart_start_col + offset), Cells(line, g_ganttchart_start_col + offset + date_len)).Select
        Selection.Merge
        ActiveCell.FormulaR1C1 = Cells(line, g_task_process_col).Value
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Style = "Percent"
        End With
        Selection.Font.Bold = True
        
        Cells(line, g_ganttchart_start_col + offset + date_len + 1).Value = Strings.Split(LTrim(Cells(line, g_task_name_col).Value), " ")(1)
        
    ElseIf Cells(line, g_task_type_col).Value = "P" Then
    
        Range(Cells(line, g_ganttchart_start_col + offset), Cells(line, g_ganttchart_start_col + offset + date_len)).Select

        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = g_GanttChartParentTaskColor
            .TintAndShade = 0
            .Weight = xlMedium
        End With

        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = g_GanttChartParentTaskColor
            .TintAndShade = 0
            .Weight = xlMedium
        End With

        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = g_GanttChartParentTaskColor
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        
        Selection.Font.Bold = True
        Selection.Font.Color = g_GanttChartParentTaskColor

        Cells(line, g_ganttchart_start_col + offset).Value = _
                Strings.Split(LTrim(Cells(line, g_task_name_col).Value), " ")(1) _
                & "(" & Cells(line, g_task_process_col).Text & ")"

    End If
    Application.CutCopyMode = False
    
End Sub

Sub drawCurrentDayLine()
    Dim offset As Integer
    Dim date_len As Integer
    Dim base As Variant
    Dim i As Integer
    
    If Not (Now() >= GetGanttMinDate() And Now <= GetGanttMaxDate()) Then
        Exit Sub
    End If
    
    base = Cells(g_start_day_x, g_start_day_y).Value
    offset = Now() - base

    Range(Cells(g_task_area_start_line, g_ganttchart_start_col + offset - 1), Cells(g_task_area_end_line, g_ganttchart_start_col + offset - 1)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone

    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    
'    With Selection.Borders(xlInsideHorizontal)
'        .LineStyle = xlContinuous
'        .ThemeColor = 5
'        .TintAndShade = 0.399945066682943
'        .Weight = xlHairline
'    End With
    
End Sub

Sub clearGanttChart()
    Dim col As Integer

    Range(Cells(g_task_area_start_line, g_ganttchart_start_col), Cells(g_task_area_end_line, g_date_end_col)).Select
    Selection.ClearContents
    Selection.Merge
    Selection.UnMerge
    
    Range("D2").Select
    Selection.Copy
    Range(Cells(g_task_area_start_line, g_ganttchart_start_col), Cells(g_task_area_end_line, g_date_end_col)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    For col = g_ganttchart_start_col To g_ganttchart_start_col + 7
        If Cells(g_task_area_start_line - 1, col).Value = "六" Then
            Exit For
        End If
    Next
    
    For col = col To g_date_end_col Step 7
        Range("L2").Select
        Selection.Copy
        Range(Cells(g_task_area_start_line, col), Cells(g_task_area_end_line, col)).Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False

        Range("N2").Select
        Selection.Copy
        Range(Cells(g_task_area_start_line, col + 1), Cells(g_task_area_end_line, col + 1)).Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False

    Next

    Application.CutCopyMode = False
    
End Sub

Function CalcWorkDate(cur_date As Date, day As Integer) As Date

End Function

Function getNextWorkDate(cur_date As Date) As Date

End Function

Function isWorkDay(cur_date As Date) As Boolean

End Function

Sub updatePrivateTasks()
    Dim line As Integer
    Dim str1 As String
    
'    For line = g_task_area_start_line To g_task_area_end_line
'        str1 = str1 & Cells(line, g_task_name_col).Value & ","
'    Next

    Range(Cells(g_task_area_start_line, g_PrivateTask_col), Cells(g_task_area_end_line, g_PrivateTask_col)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$C$6:$C$35"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
    Debug.Print "updatePrivateTasks() end"
End Sub

'update task date(Start/End)
Sub clearTaskDate(index As Integer)
    Dim i As Integer
    For i = index To g_max_level
        g_task_startday_min(i) = g_task_default_date
        g_task_endday_max(i) = g_task_default_date
    Next
End Sub

Sub clearTaskActDay(index As Integer)
    Dim i As Integer
    For i = index To g_max_level
        g_task_actday(i) = 0
    Next
End Sub

Sub clearTaskPlanDay(index As Integer)
    Dim i As Integer
    For i = index To g_max_level
        g_task_planday(i) = 0
    Next
End Sub

Sub updateParentTaskProcess(line As Integer, task_level As Integer, is_parent As Boolean)
    If is_parent = False Then
        g_task_actday(task_level) = g_task_actday(task_level) + (Cells(line, g_task_days_col).Value * Cells(line, g_task_process_col).Value)
    ElseIf is_parent = True Then
        Cells(line, g_task_process_col).Value = g_task_actday(task_level + 1) / Cells(line, g_task_days_col).Value
        
        g_task_actday(task_level) = g_task_actday(task_level) + g_task_actday(task_level + 1)
        
        Call clearTaskActDay(task_level + 1)
    End If
End Sub

Sub updateParentTaskDate(line As Integer, task_level As Integer, is_parent As Boolean)
    If is_parent = False Then
        'update task start date
        If (g_task_startday_min(task_level) = g_task_default_date) Or _
            (g_task_startday_min(task_level) > Cells(line, g_task_start_day_col).Value) Then
            g_task_startday_min(task_level) = Cells(line, g_task_start_day_col).Value
        End If
        
        'update task end date
        If (g_task_endday_max(task_level) = g_task_default_date) Or _
            (g_task_endday_max(task_level) < Cells(line, g_task_end_day_col).Value) Then
            g_task_endday_max(task_level) = Cells(line, g_task_end_day_col).Value
        End If

        g_task_planday(task_level) = g_task_planday(task_level) + Cells(line, g_task_days_col).Value
        
    ElseIf is_parent = True Then
        If (g_task_startday_min(task_level + 1) <> g_task_default_date) Then
            Cells(line, g_task_start_day_col).Value = g_task_startday_min(task_level + 1)
            Cells(line, g_task_end_day_col).Value = g_task_endday_max(task_level + 1)

            If (g_task_startday_min(task_level) = g_task_default_date) Or _
                (g_task_startday_min(task_level) > Cells(line, g_task_start_day_col).Value) Then
                g_task_startday_min(task_level) = Cells(line, g_task_start_day_col).Value
            End If
            
            If (g_task_endday_max(task_level) = g_task_default_date) Or _
                (g_task_endday_max(task_level) < Cells(line, g_task_end_day_col).Value) Then
                g_task_endday_max(task_level) = Cells(line, g_task_end_day_col).Value
            End If

            Cells(line, g_task_days_col).Value = g_task_planday(task_level + 1)
            g_task_planday(task_level) = g_task_planday(task_level) + Cells(line, g_task_days_col).Value
            
            Call clearTaskDate(task_level + 1)
            Call clearTaskPlanDay(task_level + 1)
        End If

    End If

End Sub

Sub updateTaskInfo()
    Dim line As Integer
    Dim task_level As Integer

    Call clearTaskDate(0)

    For line = g_task_area_end_line To g_task_area_start_line Step -1
        task_level = Cells(line, g_task_level_col).Value
        
        If (Cells(line, g_task_type_col).Value = "T") And (IsError(Range("G" & line)) = False) Then
        
            Call updateParentTaskDate(line, task_level, False)
            Call updateParentTaskProcess(line, task_level, False)
            
        ElseIf Cells(line, g_task_type_col).Value = "P" Then
        
            Call updateParentTaskDate(line, task_level, True)
            Call updateParentTaskProcess(line, task_level, True)
            
        End If

    Next

End Sub

Sub setBaseLineDate()
    Dim line As Integer

    For line = g_task_area_start_line To g_task_area_end_line
        Cells(line, g_task_baseline_start_day_col).Value = Cells(line, g_task_start_day_col).Value
        Cells(line, g_task_baseline_end_day_col).Value = Cells(line, g_task_end_day_col).Value
    Next

End Sub









