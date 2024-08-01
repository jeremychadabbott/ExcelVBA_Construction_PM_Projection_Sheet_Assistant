Attribute VB_Name = "WorksheetsCreate"
Sub copytabs(targetPath)
    Dim dates As Date
    Dim targetYear As Integer
    Dim targetMonth As Integer
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim tarpath As String
    Dim filename As String
    Dim Pathname As String
    Dim Sheet As Worksheet
    Dim newName As String
    Dim jobName As String
    Dim Exists As Boolean
    Dim i As Integer
    Dim latestYear As Integer
    Dim latestMonth As Integer
    Dim sheetName As String
    Dim sheetYear As Integer
    Dim sheetMonth As Integer
    Dim Sheetdate As Date
    Dim response As VbMsgBoxResult
    Dim ws As Workbook

    ' Get target date from cell L4
    dates = ThisWorkbook.Sheets(1).Range("L4").Value
    targetYear = Year(dates)
    targetMonth = Month(dates)
    
    tarpath = targetPath & "\Projection Sheets"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(tarpath)

    Application.ScreenUpdating = False

    ' Initialize latestYear and latestMonth to track the most recent sheet
    latestYear = 0
    latestMonth = 0

    ' Loop through files in the target folder
    For Each objFile In objFolder.Files
        If objFile.Name Like "*Projections*" And Not objFile.Name Like "*$*" Then
            filename = objFile.Name
            Pathname = objFile.Path

            ' Open the workbook and process sheets
            Set ws = Workbooks.Open(Pathname)

            ' Check if target month's sheets already exist
            For Each Sheet In ws.Sheets
                If Sheet.Name Like "* " & targetYear & "-" & Format(targetMonth, "00") & "*" Then
                    Debug.Print "Sheet for target month already exists: " & Sheet.Name
                    ws.Close SaveChanges:=False
                    Application.ScreenUpdating = True
                    Exit Sub ' Sheets for target month already exist, exit the subroutine
                End If
            Next Sheet

            ' Find the most recent sheet available
            For Each Sheet In ws.Sheets
                sheetName = Sheet.Name
                ' Extract the date from the sheet name
                If sheetName Like "*[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]" And Not LCase(sheetName) Like "*qtr*" And Not LCase(sheetName) Like "*([0-9])*" Then
                    'Debug.Print "Checking Sheet Name: " & sheetName
                    Sheetdate = Right(sheetName, 10)
                    Sheetdate = Replace(Sheetdate, " ", "")
                    Sheetdate = CDate(Sheetdate)
                    sheetYear = Year(Sheetdate)
                    sheetMonth = Month(Sheetdate)
                    'Debug.Print "Parsed Year: " & sheetYear & ", Parsed Month: " & sheetMonth
                    'MsgBox "Parsed Year: " & sheetYear & ", Parsed Month: " & sheetMonth
                    If sheetYear > latestYear Or (sheetYear = latestYear And sheetMonth > latestMonth) Then
                        latestYear = sheetYear
                        latestMonth = sheetMonth
                        Debug.Print "Latest Sheet Year is now->" & latestYear & " " & latestMonth
                    End If
                End If
            Next Sheet

            Debug.Print "Latest Year: " & latestYear
            Debug.Print "Latest Month: " & latestMonth


            ' Ensure the latest date is not beyond the target date
            If latestYear > targetYear Or (latestYear = targetYear And latestMonth > targetMonth) Then
                MsgBox "Error, when creating new sheets, there is newer sheets than the target date in cell 4 of the projection sheet helper"
            End If

            ' Process sheets from the identified last month and year
            For Each Sheet In ws.Sheets
                'Debug.Print "Processing Sheet: " & Sheet.Name
                If Sheet.Name Like "* " & latestYear & "-" & Format(latestMonth, "00") & "*" Then
                    jobName = Left(Sheet.Name, 5)
                    newName = jobName & " " & targetYear & "-" & Format(targetMonth, "00") & "-" & Day(dates)
                    For Repeat = 1 To 5
                        newName = Replace(newName, "  ", " ")
                    Next Repeat
            
                    'Copying forward sheet (if needed)
                    Debug.Print "Copying forward sheet:" & Sheet.Name
                    
                    
                    ' Check if proposed new tab already exists
                    Exists = False
                    For i = 1 To ws.Sheets.count
                        If ws.Sheets(i).Name = newName Then
                            Exists = True
                        End If
                    Next i

                    Debug.Print "New Name: " & newName
                    Debug.Print "Exists: " & Exists

                    ' Copy forward tab
                    If Not Exists Then
                        ws.Sheets(Sheet.Name).Copy After:=ws.Sheets(ws.Sheets.count)
                        With ActiveSheet
                            .Name = newName
                            .Tab.ColorIndex = Month(dates)
                        End With
                    End If
                End If
            Next Sheet

            ' Populate the sheet with data
            Call SageReportCheck(targetPath)
            For Each Sheet In ws.Sheets
                Dim TempYears As Integer
                TempYears = targetYear
                If Sheet.Name Like "*" & TempYears & "-" & Format(targetMonth, "00") & "*" Then
                    newName = Sheet.Name
                    Call transferdata(targetPath, filename, newName)
                End If
            Next Sheet

            ws.Close SaveChanges:=True
        End If
    Next objFile

    Application.ScreenUpdating = True
End Sub

