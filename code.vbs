Option Explicit

Const g_max_level As Integer = 5
Dim g_level_number(g_max_level) As Integer

Const g_level1_format As Integer = 0
Const g_level2_format As Integer = 2
Const g_level3_format As Integer = 4
Const g_level4_format As Integer = 6
Const g_level5_format As Integer = 8

Const g_start_day_x As Integer = 3
Const g_start_day_y As Integer = 18
Const g_ganttchart_start_col As Integer = 18

Const g_task_start_line As Integer = 6
Dim g_task_end_line As Integer

Dim g_date_end_col As Integer

Sub test()
    MsgBox DateAdd("m", 6, Cells(g_start_day_x, g_start_day_y).Value)
End Sub


Sub getTaskArea()
    Dim line As Integer
    Dim col As Integer

    line = g_task_start_line
    Do While Cells(line, 3).Value <> ""
        line = line + 1
    Loop
    g_task_end_line = line
    
    
    g_date_end_col = g_ganttchart_start_col + DateAdd("m", 6, Cells(g_start_day_x, g_start_day_y).Value) - Cells(g_start_day_x, g_start_day_y).Value
    
    For col = g_start_day_y + 1 To g_date_end_col
        Cells(g_start_day_x + 1, col).Value = Cells(g_start_day_x + 1, col - 1).Value + 1
    Next
    
End Sub

Sub UpdateGanttChart()
    Dim line As Integer
    Dim level As Integer
    
    Application.ScreenUpdating = False
    
    Call getTaskArea
    
    Call clearLevelNumber
    Call clearGanttChart
    
    For line = g_task_start_line To g_task_end_line
        formatLevel (line)
        drawGanttChart (line)
    Next
    g_task_end_line = line
    
    Call drawCurrentDayLine
    Application.ScreenUpdating = True
    
End Sub

Sub formatLevel(line As Integer)
    Dim level As Integer
    
    level = checkLevel(Cells(line, 3).Value)
    LevelNumberAdd (level)
    If InStr(LTrim(Cells(line, 3).Value), " ") Then
        Cells(line, 3).Value = getLevelHead(level) & getLevel(level) & " " & Strings.Split(LTrim(Cells(line, 3).Value), " ")(1)
    Else
        Cells(line, 3).Value = getLevelHead(level) & getLevel(level) & " " & LTrim(Cells(line, 3).Value)
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

Function getLevel(level As Integer) As String
    Dim i As Integer
    level = level - 1
    If level > g_max_level Then
        MsgBox "Level" & level & "Error"
    End If
    
    getLevel = getLevel & g_level_number(i)
    For i = 1 To level
        getLevel = getLevel & "." & g_level_number(i)
    Next
End Function

Function getLevelHead(level As Integer) As String
    If level > g_max_level Then
        MsgBox "Level" & level & "Error"
    End If
    
    Select Case level
        Case 1
            getLevelHead = Space(0)
        Case 2
            getLevelHead = Space(g_level2_format)
        Case 3
            getLevelHead = Space(g_level3_format)
        Case 4
            getLevelHead = Space(g_level4_format)
        Case 5
            getLevelHead = Space(g_level5_format)
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
    
    base = Cells(g_start_day_x, g_start_day_y).Value
    date_len = Cells(line, 8).Value - Cells(line, 7).Value
    offset = Cells(line, 7).Value - base
    
    If Cells(line, 5).Value = 1 Then
        Range("F2").Select
    ElseIf Cells(line, 5).Value > 0 Then
        Range("J2").Select
    Else
        Range("H2").Select
    End If
    Selection.Copy
    
    For i = 0 To date_len
        Cells(line, g_ganttchart_start_col + offset + i).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                                  SkipBlanks:=False, Transpose:=False
    Next
    
    Application.CutCopyMode = False
    
End Sub

Sub drawCurrentDayLine()
    Dim offset As Integer
    Dim date_len As Integer
    Dim base As Variant
    Dim i As Integer
    
    base = Cells(g_start_day_x, g_start_day_y).Value
    offset = Now() - base

    Range(Cells(g_task_start_line, g_ganttchart_start_col + offset - 1), Cells(g_task_end_line, g_ganttchart_start_col + offset - 1)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone

    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -16776961
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 5
        .TintAndShade = 0.399945066682943
        .Weight = xlHairline
    End With
    
End Sub

Sub clearGanttChart()

    Range("D2").Select
    
    Selection.Copy
    Range("R6:FR40").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
    Application.CutCopyMode = False
    
End Sub

Function CalcWorkDate(cur_date As Date, day As Integer) As Date

End Function

Function getNextWorkDate(cur_date As Date) As Date

End Function

Function isWorkDay(cur_date As Date) As Boolean

End Function
