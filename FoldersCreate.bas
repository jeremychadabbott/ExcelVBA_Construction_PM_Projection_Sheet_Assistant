Attribute VB_Name = "FoldersCreate"
Sub foldercheck(targetPath)
    Dim dates As Date
    Dim Years As Integer
    Dim months As Integer
    Dim Found As Integer
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objSubFolder As Object
    Dim tarfolder As String
    Dim tarpath As String
    Dim Mon As String
    Dim Quarter As String
    
    dates = ThisWorkbook.Sheets(1).Range("L4").Value
    Years = Year(dates)
    months = Month(dates)
    
    ' Initialize FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    targetPath = ThisWorkbook.Path
    Set objFolder = objFSO.GetFolder(targetPath)
    
    ' Check for Year Folder
    Found = 0
    For Each objSubFolder In objFolder.subfolders
        tarfolder = UCase(Replace(Replace(objSubFolder.Name, " ", ""), "-", ""))
        If tarfolder Like "*" & Years & "*" Then
            Found = 1
            tarpath = objSubFolder.Path
            Exit For
        End If
    Next objSubFolder
    
    If Found = 0 Then
        tarpath = targetPath & "\" & Years
        MkDir tarpath
    End If
    
    ' Check for Quarter Folder
    Found = 0
    targetPath = tarpath
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
    
    Set objFolder = objFSO.GetFolder(targetPath)
    For Each objSubFolder In objFolder.subfolders
        tarfolder = UCase(Replace(Replace(objSubFolder.Name, " ", ""), "-", ""))
        If tarfolder Like "*" & Quarter & "*" Then
            Found = 1
            tarpath = objSubFolder.Path
            Exit For
        End If
    Next objSubFolder
    
    If Found = 0 Then
        tarpath = targetPath & "\" & Quarter
        If Dir(tarpath) <> "" Then MkDir tarpath
    End If
    
    ' Check for Month Folder
    Found = 0
    targetPath = tarpath
    Set objFolder = objFSO.GetFolder(targetPath)
    For Each objSubFolder In objFolder.subfolders
        tarfolder = UCase(Replace(Replace(objSubFolder.Name, " ", ""), "-", ""))
        If tarfolder Like "*" & Mon & "*" Then
            Found = 1
            tarpath = objSubFolder.Path
            Exit For
        End If
    Next objSubFolder
    
    If Found = 0 Then
        tarpath = targetPath & "\" & Mon
        If Dir(tarpath) <> "" Then MkDir tarpath
    End If
    
    ' Check if subfolders exist and create if they don't
    Dim subfolders As Variant
    subfolders = Array("Backup Reports", "Bank Statements", "Financial Reports", "Projection Sheets", "Schedules")
    
    For Each folder In subfolders
        If Dir(tarpath & "\" & folder, vbDirectory) = "" Then MkDir tarpath & "\" & folder
    Next folder
    
    targetPath = tarpath
End Sub

