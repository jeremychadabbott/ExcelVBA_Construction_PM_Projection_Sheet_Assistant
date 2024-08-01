Attribute VB_Name = "Start"
Sub Button1_Click()
    Dim remote As Variant
    Dim targetPath As String

    remote = ThisWorkbook.Sheets(1).Range("E1").Value
    Application.ScreenUpdating = False

    ' Check if folders need to be created
    Call foldercheck(targetPath) ' Module 2

    ' Check if projection sheets need to be created
    Call projectionsheet(targetPath) ' Module 3

    ' Create new tabs in projection sheet if needed
    Call copytabs(targetPath) ' Module 5

    Application.ScreenUpdating = True

    ' Update projection sheets with information from reports
    Call SageReportCheck(targetPath)

    UserForm1.Hide

    If remote <> 1 Then
        MsgBox "Finished!"
    End If
End Sub

