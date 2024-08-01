Attribute VB_Name = "SageReports_FetchData"
Sub SageReportCheck(targetPath)
    Dim objFSO As Object
    Dim objFolder As Object
    Dim reports As Integer
    Dim locatePath As String
    
    Application.ScreenUpdating = False
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(targetPath & "\Backup Reports")
    
    reports = CountRequiredReports(objFolder)
    
    If reports < 3 Then
        MsgBox "Could not find all the Sage reports needed in the folder " & targetPath & _
               ", looking to verify the existence of *Committed Costs*, *Job Labor Totals* & *Over Under Billings*. " & _
               "Check these reports are in the backup reports folder and named correctly. Assistant will Exit now"
        Exit Sub
    End If
    
    locatePath = targetPath & "\Backup Reports\Committed Costs.xlsx"
    ProcessCommittedCostsReport locatePath
    
    Application.ScreenUpdating = True
End Sub

Function CountRequiredReports(folder As Object) As Integer
    Dim file As Object
    Dim count As Integer
    
    For Each file In folder.Files
        If InStr(file.Name, "Committed Costs") > 0 Then count = count + 1
        If InStr(file.Name, "Job Labor Totals") > 0 Then count = count + 1
        If InStr(file.Name, "Over Under Billings") > 0 Then count = count + 1
    Next file
    
    CountRequiredReports = count
End Function

Sub ProcessCommittedCostsReport(filePath As String)
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = Workbooks.Open(filePath)
    
    For Each ws In wb.Worksheets
        ProcessWorksheet ws
    Next ws
    
    wb.Close SaveChanges:=False
End Sub

Sub ProcessWorksheet(ws As Worksheet)
    Dim jobName As String
    Dim y As Long
    
    jobName = FindJobName(ws)
    
    If jobName = "" Then Exit Sub
    
    RenameWorksheet ws, jobName
    AddProjectionSection ws, y
End Sub

Function FindJobName(ws As Worksheet) As String
    Dim x As Long, y As Long
    Dim cellValue As String
    
    For x = 0 To 1
        For y = 0 To 8
            cellValue = ws.Range("A1").Offset(y, x).Value
            If cellValue Like "*Job*" Then
                FindJobName = FormatJobName(cellValue)
                Exit Function
            End If
        Next y
    Next x
    
    FindJobName = ""
End Function

Function FormatJobName(jobName As String) As String
    jobName = Replace(jobName, Left(jobName, 5), "")
    
    If jobName = "" Then
        MsgBox "ERROR, while reformatting the jobname somehow the jobname was deleted and equals nothing. Cannot rename tab to nothing"
        Exit Function
    End If
    
    If jobName Like "######*" Then jobName = Replace(jobName, Left(jobName, 7), "")
    
    If jobName Like "EC*" Or jobName Like "DD*" Or jobName Like "GR*" Or jobName Like "*WARRANTY*" Then
        jobName = Left(Replace(jobName, " ", ""), 7)
    Else
        jobName = Left(jobName, 4)
    End If
    
    FormatJobName = jobName
End Function

Sub RenameWorksheet(ws As Worksheet, newName As String)
    Dim m As Long
    Dim Found As Boolean
    
    If ws.Name <> newName Then
        Found = False
        For m = 1 To ws.Parent.Worksheets.count
            If ws.Parent.Worksheets(m).Name = newName Then
                Found = True
                Exit For
            End If
        Next m
        
        If Not Found Then ws.Name = newName
    End If
End Sub

Sub AddProjectionSection(ws As Worksheet, y As Long)
    SetupProjectionColumns ws, y
    AddSubTotalsFormulas ws, y
    AddGrandTotalsFormulas ws, y
End Sub

Sub SetupProjectionColumns(ws As Worksheet, y As Long)
    ' ROW Q, Job cost plus committed
    SetupColumnQ ws, y
    
    ' ROW R, percent spent against budget
    SetupColumnR ws, y
    
    ' ROW S, Computed Final Cost
    SetupColumnS ws, y
    
    ' ROW T, PM Override % Complete
    SetupColumnT ws, y
    
    ' ROW U, Adjusted Final Cost
    SetupColumnU ws, y
End Sub

Sub SetupColumnQ(ws As Worksheet, y As Long)
    With ws
        .Range("P4").Offset(y) = "*Entered by Automation"
        .Range("Q4").Offset(y).Font.Bold = True
        .Range("Q4").Offset(y + 1) = "Committed"
        .Range("Q4").Offset(y + 1).Interior.ColorIndex = 35
        .Range("Q4").Offset(y + 1).Font.Bold = True
        .Range("Q4").Offset(y + 2) = "Remaining"
        .Range("Q4").Offset(y + 2).Interior.ColorIndex = 35
        .Range("Q4").Offset(y + 2).Font.Bold = True
        .Range("Q4").Offset(y + 3) = "    +"
        .Range("Q4").Offset(y + 3).Interior.ColorIndex = 35
        .Range("Q4").Offset(y + 4) = "Cost to Date"
        .Range("Q4").Offset(y + 4).Font.Bold = True
        .Range("Q4").Offset(y + 4).Interior.ColorIndex = 35
    End With
End Sub

Sub SetupColumnR(ws As Worksheet, y As Long)
    With ws
        .Range("R4").Offset(y + 3) = "    %"
        .Range("R4").Offset(y + 3).Interior.ColorIndex = 35
        .Range("R4").Offset(y + 4) = "Complete"
        .Range("R4").Offset(y + 4).Interior.ColorIndex = 35
        .Range("R4").Offset(y + 4).Font.Bold = True
    End With
End Sub

Sub SetupColumnS(ws As Worksheet, y As Long)
    With ws
        .Range("S4").Offset(y + 3) = "Computed"
        .Range("S4").Offset(y + 3).Interior.ColorIndex = 35
        .Range("S4").Offset(y + 4) = "Final Cost"
        .Range("S4").Offset(y + 4).Interior.ColorIndex = 35
        .Range("S4").Offset(y + 4).Font.Bold = True
    End With
End Sub

Sub SetupColumnT(ws As Worksheet, y As Long)
    With ws
        .Range("T4").Offset(y + 1) = "PM"
        .Range("T4").Offset(y + 1).Interior.ColorIndex = 35
        .Range("T4").Offset(y + 1).Font.Bold = True
        .Range("T4").Offset(y + 2) = "Override"
        .Range("T4").Offset(y + 2).Interior.ColorIndex = 35
        .Range("T4").Offset(y + 2).Font.Bold = True
        .Range("T4").Offset(y + 3) = "    %"
        .Range("T4").Offset(y + 3).Interior.ColorIndex = 35
        .Range("T4").Offset(y + 4) = "Complete"
        .Range("T4").Offset(y + 4).Font.Bold = True
        .Range("T4").Offset(y + 4).Interior.ColorIndex = 35
    End With
End Sub

Sub SetupColumnU(ws As Worksheet, y As Long)
    With ws
        .Range("U4").Offset(y + 3) = "Adjusted"
        .Range("U4").Offset(y + 3).Interior.ColorIndex = 35
        .Range("U4").Offset(y + 4) = "Final Cost"
        .Range("U4").Offset(y + 4).Font.Bold = True
        .Range("U4").Offset(y + 4).Interior.ColorIndex = 35
    End With
End Sub

Sub AddSubTotalsFormulas(ws As Worksheet, y As Long)
    Dim x As Long
    
    For x = y To 1000
        If ws.Range("D4").Offset(x) = "Sub Totals:" Then
            AddSubTotalFormulas ws, x + 4
            Exit For
        End If
    Next x
    
    ws.Range("Q:U").EntireColumn.AutoFit
End Sub

Sub AddSubTotalFormulas(ws As Worksheet, row As Long)
    With ws
        .Range("Q" & row).Formula = "=+M" & row & "+J" & row
        .Range("Q" & row).Font.Bold = True
        .Range("Q" & row).Interior.ColorIndex = 35
        
        .Range("R" & row).Formula = "=+Q" & row & "/F" & row
        .Range("R" & row).Font.Bold = True
        .Range("R" & row).NumberFormat = "0%"
        .Range("R" & row).Interior.ColorIndex = 35
        
        .Range("S" & row).Formula = "=+F" & row
        .Range("S" & row).Font.Bold = True
        .Range("S" & row).Interior.ColorIndex = 35
        
        .Range("T" & row).NumberFormat = "0%"
        .Range("T" & row).Font.Bold = True
        .Range("T" & row).Formula = "=R" & row
        .Range("T" & row).Interior.ColorIndex = 35
        
        .Range("U" & row).Formula = "=+Q" & row & "/T" & row
        .Range("U" & row).Font.Bold = True
        .Range("U" & row).NumberFormat = "#,##0.00"
        .Range("U" & row).Interior.ColorIndex = 35
    End With
End Sub

Sub AddGrandTotalsFormulas(ws As Worksheet, y As Long)
    Dim x As Long
    
    For x = y To 1000
        If ws.Range("D4").Offset(x) = "Grand Totals:" Then
            AddGrandTotalFormulas ws, x + 4
            Exit For
        End If
    Next x
End Sub

Sub AddGrandTotalFormulas(ws As Worksheet, row As Long)
    With ws
        .Range("Q" & row).Formula = "=SUM(Q12:Q" & (row - 1) & ")"
        .Range("Q" & row).NumberFormat = "#,##0.00"
        .Range("Q" & row).Font.Bold = True
        .Range("Q" & row).Interior.ColorIndex = 35
        
        .Range("S" & row).Formula = "=SUM(S12:S" & (row - 1) & ")"
        .Range("S" & row).NumberFormat = "#,##0.00"
        .Range("S" & row).Font.Bold = True
        .Range("S" & row).Interior.ColorIndex = 35
        
        .Range("U" & row).Formula = "=SUM(U12:U" & (row - 1) & ")"
        .Range("U" & row).NumberFormat = "#,##0.00"
        .Range("U" & row).Font.Bold = True
        .Range("U" & row).Interior.ColorIndex = 35
        
        .Range("T" & row + 2) = .Range("D" & row + 2)
        .Range("T" & row + 2).HorizontalAlignment = xlRight
        .Range("T" & row + 2).Font.Bold = True
        
        .Range("T" & row + 3) = .Range("D" & row + 3)
        .Range("T" & row + 3).HorizontalAlignment = xlRight
        .Range("T" & row + 3).Font.Bold = True
        
        .Range("T" & row + 4) = .Range("D" & row + 4)
        .Range("T" & row + 4).HorizontalAlignment = xlRight
        .Range("T" & row + 4).Font.Bold = True
        
        .Range("U" & row + 2) = .Range("F" & row + 2)
        .Range("U" & row + 2).NumberFormat = "#,##0.00"
        .Range("U" & row + 2).Font.Bold = True
        
        .Range("U" & row + 3) = .Range("F" & row + 3)
        .Range("U" & row + 3).NumberFormat = "#,##0.00"
        .Range("U" & row + 3).Font.Bold = True
        
        .Range("U" & row + 4).Formula = "=+U" & (row + 2) & "-U" & row
        .Range("U" & row + 4).NumberFormat = "#,##0.00"
        .Range("U" & row + 4).Font.Bold = True
        .Range("U" & row + 4).Interior.ColorIndex = 35
    End With
End Sub

