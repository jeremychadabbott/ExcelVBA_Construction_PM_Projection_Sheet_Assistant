Attribute VB_Name = "workbooksCreate"
Sub projectionsheet(targetPath)
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim tarpath As String
    Dim dates As Date
    Dim Years As Integer
    Dim months As Integer
    Dim Mon As String
    Dim Quarter As String
    Dim oldtarpath As String
    Dim originalname As String
    Dim originalpath As String
    Dim newName As String
    Dim leftnewname As String
    Dim rightnewname As String
    Dim allcount As Integer
    Dim count As Integer
    
    tarpath = targetPath & "\Projection Sheets"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(tarpath)
    
    ' Check if projection sheets already exist in the folder
    For Each objFile In objFolder.Files
        If InStr(objFile.Name, "Projections") > 0 Then Exit Sub
    Next objFile
    
    ' Determine last month's directory
    dates = ThisWorkbook.Sheets(1).Range("L4").Value
    Years = Year(dates)
    months = Month(dates) - 1
    
    If months = 0 Then
        months = 12
        Years = Years - 1
    End If
    
    Select Case months
        Case 1
            Mon = "01-January " & Years
            Quarter = "1st Qtr " & Years
        Case 2
            Mon = "02-February " & Years
            Quarter = "1st Qtr " & Years
        Case 3
            Mon = "03-March " & Years
            Quarter = "1st Qtr " & Years
        Case 4
            Mon = "04-April " & Years
            Quarter = "2nd Qtr " & Years
        Case 5
            Mon = "05-May " & Years
            Quarter = "2nd Qtr " & Years
        Case 6
            Mon = "06-June " & Years
            Quarter = "2nd Qtr " & Years
        Case 7
            Mon = "07-July " & Years
            Quarter = "3rd Qtr " & Years
        Case 8
            Mon = "08-August " & Years
            Quarter = "3rd Qtr " & Years
        Case 9
            Mon = "09-September " & Years
            Quarter = "3rd Qtr " & Years
        Case 10
            Mon = "10-October " & Years
            Quarter = "4th Qtr " & Years
        Case 11
            Mon = "11-November " & Years
            Quarter = "4th Qtr " & Years
        Case 12
            Mon = "12-December " & Years
            Quarter = "4th Qtr " & Years
    End Select
    
    oldtarpath = ThisWorkbook.Path & "\" & Years & "\" & Quarter & "\" & Mon & "\Projection Sheets"
    
    If Dir(oldtarpath, vbDirectory) = "" Then
        MsgBox "Program can't find the folder 'Projection Sheets' in last month's folder location. The computer tried to find it here: " & oldtarpath
        Dim fn As Variant
        fn = Application.GetOpenFilename("EXCEL FILES,*.XLSX", 1, "Select location of XLSX to import", , False)
        If TypeName(fn) = "Boolean" Then Exit Sub
        oldtarpath = Left(fn, InStrRev(fn, "\") - 1)
    End If
    
    Set objFolder = objFSO.GetFolder(oldtarpath)
    
    ' Copy and rename projection sheets
    For Each objFile In objFolder.Files
        If InStr(objFile.Name, "Projections") > 0 Then
            originalname = objFile.Name
            originalpath = objFile.Path
            newName = Replace(objFile.Name, "January", "February")
            newName = Replace(newName, "February", "March")
            newName = Replace(newName, "March", "April")
            newName = Replace(newName, "April", "May")
            newName = Replace(newName, "May", "June")
            newName = Replace(newName, "June", "July")
            newName = Replace(newName, "July", "August")
            newName = Replace(newName, "August", "September")
            newName = Replace(newName, "September", "October")
            newName = Replace(newName, "October", "November")
            newName = Replace(newName, "November", "December")
            newName = Replace(newName, "December", "January")
            
            allcount = Len(newName)
            count = InStr(1, newName, "20") - 1
            If newName Like "*January*" Then
                leftnewname = Left(newName, count) & (Years + 1)
            Else
                leftnewname = Left(newName, count) & Years
            End If
            rightnewname = Right(newName, allcount - count - 4)
            newName = leftnewname & rightnewname
            
            If Dir(tarpath & "\" & newName) = "" Then
                FileCopy originalpath, tarpath & "\" & originalname
                Name tarpath & "\" & originalname As tarpath & "\" & newName
            Else
                If ThisWorkbook.Sheets(1).Range("E1").Value <> 1 Then _
                    MsgBox "Projection sheet " & newName & " already exists in the target folder. Continuing will replace existing data with new and clear out the projections."
            End If
        End If
    Next objFile
End Sub


