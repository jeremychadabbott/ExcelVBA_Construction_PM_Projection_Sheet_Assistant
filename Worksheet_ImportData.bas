Attribute VB_Name = "Worksheet_ImportData"


Sub transferdata(targetPath, filename, newName)

'MsgBox "Made it to transfer Data module, " & Chr(13) & "targetpath->" & targetpath & Chr(13) & "filename->" & filename & _
    Chr(13) & "newname->" & newname

Application.ScreenUpdating = False
            DoEvents
            UserForm1.Label1.Caption = "importing data to: " & newName & " in workbook " & Workbooks(filename).Name
            UserForm1.Show vbModeless
            'UserForm1.Label1.Caption = "creating sheet and importing data: " & ActiveSheet.Name & " in workbook " & ActiveWorkbook.Name
            DoEvents
            
'Data to transfer onto projection sheets:
'                    ActiveSheet.Range("H4") = Date
'                    ActiveSheet.Range("D9") = Date
'                    ActiveSheet.Range("G8") = 0 'Labor
'                    ActiveSheet.Range("G16") = 0 'Material Invoices
'                    ActiveSheet.Range("G18") = 0 'Equipment
'                    ActiveSheet.Range("G20") = 0 'Subcontractor
'                    ActiveSheet.Range("G22") = 0 'Expenses/bonds/permits
'                    ActiveSheet.Range("G35") = 0 'Committed Material
'                    ActiveSheet.Range("G36") = 0 'Committed Equipment
'                    ActiveSheet.Range("G37") = 0 'Committed Subcontractor
'                    ActiveSheet.Range("G38") = 0 'Committed Other
'
'                    ActiveSheet.Range("G40") = 0 'Estimated Material to Buy
'                    ActiveSheet.Range("G42") = 0 'Estimated Labor Burden
'                    ActiveSheet.Range("G44") = 0 'Estimated Equipment/Subcontract'
'                    ActiveSheet.Range("G46") = 0 'Expenses
'                    ActiveSheet.Range("G48") = 0 'IBEW Work recovery
'
'                    ActiveSheet.Range("G85") = 0 'Invoiced to Date
'                    ActiveSheet.Range("G91") = 0 'Cash Collected

'Sheet Committed Costs     -> Committed, budgets and costs
'Sheet Job Labor Totals    -> Job Hours
'Sheet Over-Under Billings -> Thought I could get invoiced to date but that has tax in it
'----------------------------------
'Stage 1: Get Committed Costs
'-----------------------------------
locatePath = targetPath & "\Backup Reports\Committed Costs.xlsx"


Application.ScreenUpdating = False
Workbooks.Open (locatePath)
    For Each Sheet In Workbooks("Committed Costs.xlsx").Sheets
    
        For y = 0 To 4
        jobName = Sheet.Range("A4").Offset(y, 0)
            If jobName Like "*[0-9][0-9][0-9][0-9]*" Then
                Found = 1
                For Repeat = 1 To Len(jobName) - 4
                    If Mid(jobName, Repeat, 4) Like "[0-9][0-9][0-9][0-9]" Then
                        jobName = Mid(jobName, Repeat, 4)
                        Exit For
                    End If
                Next Repeat
                Exit For
            End If
        Next y
        
        'Reformat Committed Cost Report
        
        'MsgBox "searching committed cost report for matching sheet " & Left(newname, 4) & " & " & jobname
            If jobName Like "*" & Left(newName, 4) & "*" Then
                'MsgBox "matched up committed cost report info to " & jobname
                'Rename Sheet
                Sheet.Name = Left(newName, 4) 'if error is here, it's because there's already a sheet by that name
                '------TRANSFER DATA---------
                'material invoices------------------------------------------------------
                For x = 0 To 1000
                    If Sheet.Range("B1").Offset(x) = "Material" Then
                        'MsgBox "first target locked!"
                        For y = x To (x + 200)
                            If Sheet.Range("e1").Offset(y) = "Subtotals:" Then
                             'MsgBox "second target locked!"
                             'MsgBox "Committed material for " & Left(newname, 4) & " is " & Sheet.Range("h1").Offset(y)
                             'Committed Material
                             Workbooks(filename).Sheets(newName).Range("G35") = Sheet.Range("g1").Offset(y)
                             'MsgBox "new committed material data is:" & Workbooks(filename).Sheets(newname).Range("G35")
                             'Material Costs to Date
                             Workbooks(filename).Sheets(newName).Range("G16") = Sheet.Range("i1").Offset(y)
                             'Material Budget
                             Workbooks(filename).Sheets(newName).Range("I6") = "Material Budget"
                             Workbooks(filename).Sheets(newName).Range("J6") = Sheet.Range("F1").Offset(y)
                            Exit For
                            Else: End If
                        Next y
                        Exit For
                    Else: End If
                Next x
                
                'Labor--------------------------------------------------------------
                For x = 0 To 1000
                    If Sheet.Range("B1").Offset(x) = "Labor" Then
                        'MsgBox "first target locked!"
                        For y = x To (x + 200)
                            If Sheet.Range("e1").Offset(y) = "Subtotals:" Then
                             'MsgBox "second target locked!"
                             
                             'Labor Costs to Date
                             Workbooks(filename).Sheets(newName).Range("G8") = Sheet.Range("i1").Offset(y)
                             'Labor Budget
                             Workbooks(filename).Sheets(newName).Range("I7") = "Labor Budget"
                             Workbooks(filename).Sheets(newName).Range("J7") = Sheet.Range("F1").Offset(y)
                            Exit For
                            Else: End If
                        Next y
                        Exit For
                    Else: End If
                Next x
                'Equipment --------------------------------------------------------------
                For x = 0 To 1000
                    If Sheet.Range("B1").Offset(x) = "Equipment" Then
                        'MsgBox "first target locked!"
                        For y = x To (x + 200)
                            If Sheet.Range("e1").Offset(y) = "Subtotals:" Then
                             'MsgBox "second target locked!"
                            'Committed Equipment
                             Workbooks(filename).Sheets(newName).Range("G36") = Sheet.Range("g1").Offset(y)
                             'Equipment Costs to Date
                             Workbooks(filename).Sheets(newName).Range("G18") = Sheet.Range("i1").Offset(y)
                             'Equipment Budget
                             Workbooks(filename).Sheets(newName).Range("I8") = "Equipment Budget"
                             Workbooks(filename).Sheets(newName).Range("J8") = Sheet.Range("F1").Offset(y)
                            Exit For
                            Else: End If
                        Next y
                        Exit For
                    Else: End If
                Next x
                'Subcontractor --------------------------------------------------------------
                For x = 0 To 1000
                    If Sheet.Range("B1").Offset(x) = "Subcontractor" Then
                        'MsgBox "first target locked!"
                        For y = x To (x + 200)
                            If Sheet.Range("e1").Offset(y) = "Subtotals:" Then
                             'MsgBox "second target locked!"
                            'Committed Subcontractor
                             Workbooks(filename).Sheets(newName).Range("G37") = Sheet.Range("g1").Offset(y)
                             'Subcontractor Costs to Date
                             Workbooks(filename).Sheets(newName).Range("G20") = Sheet.Range("i1").Offset(y)
                             'Subcontractor Budget
                             Workbooks(filename).Sheets(newName).Range("I9") = "Subcontractor Budget"
                             Workbooks(filename).Sheets(newName).Range("J9") = Sheet.Range("F1").Offset(y)
                            Exit For
                            Else: End If
                        Next y
                        Exit For
                    Else: End If
                Next x
                
                'Other ------------------------------------------------
                For x = 0 To 1000
                    If Sheet.Range("B1").Offset(x) = "Other" Then
                        'MsgBox "first target locked!"
                        For y = x To (x + 200)
                            If Sheet.Range("e1").Offset(y) = "Subtotals:" Then
                             'MsgBox "second target locked!"
                            'Other Committed
                             Workbooks(filename).Sheets(newName).Range("G38") = Sheet.Range("g1").Offset(y)
                             'Other Costs to Date
                             Workbooks(filename).Sheets(newName).Range("G22") = Sheet.Range("i1").Offset(y)
                             'Other Budget
                             Workbooks(filename).Sheets(newName).Range("I10") = "*Other Budget"
                             Workbooks(filename).Sheets(newName).Range("J10") = Sheet.Range("F1").Offset(y)
                            Exit For
                            Else: End If
                        Next y
                        Exit For
                    Else: End If
                Next x
                'Change Orders ------------------------------------------------
                For x = 0 To 1000
                    If Sheet.Range("D1").Offset(x) = "Actual Selling Price:" Then
                        'MsgBox "first target locked!"
                        'Workbooks(filename).Sheets(newname).Range("I60") = "*Change orders auto calculate done by subtracting G58 from *actual Sell* on committed cost report"
                        'MsgBox "*actual sell = " & Sheet.Range("F1").Offset(x)
                        'MsgBox "*contract is = " & Workbooks(filename).Sheets(newname).Range("G58")
                        'MsgBox "*result is = " & (Sheet.Range("F1").Offset(x) - Workbooks(filename).Sheets(newname).Range("G58"))
                        'Workbooks(filename).Sheets(newname).Range("G60") = ((Workbooks(filename).Sheets(newname).Range("G58")) - (Sheet.Range("F1").Offset(y)))
                        'Workbooks(filename).Sheets(newname).Range("G60") = (Sheet.Range("F1").Offset(x) - Workbooks(filename).Sheets(newname).Range("G58"))
                    Exit For
                    Else: End If
                Next x
                  
                Workbooks(filename).Sheets(newName).Range("H26").Formula = "=SUM(G8:G24)"
                Workbooks(filename).Sheets(newName).Range("H31").Formula = "=SUM(H26)"
                
     'end of committed costs sheet uses
            Else: End If
    Next Sheet
    
Workbooks("Committed Costs.xlsx").Saved = True
Workbooks("Committed Costs.xlsx").Close


'----------------------------------
'Stage 2: Get job labor totals (hours)
'-----------------------------------
locatePath = targetPath & "\Backup Reports\Job Labor Totals.xlsx"
Application.ScreenUpdating = False
Workbooks.Open (locatePath)

    'Labor Hours ------------------------------------------------
    For x = 0 To 1000
        If Sheets(1).Range("A6").Offset(x) Like "*" & Left(newName, 4) & "*" Then
            'MsgBox "first target locked!"
            'job labor hours
            Workbooks(filename).Sheets(newName).Range("C8") = Workbooks("Job Labor Totals.xlsx").Sheets(1).Range("F6").Offset(x)
            'job labor cost
            Workbooks(filename).Sheets(newName).Range("G8") = Workbooks("Job Labor Totals.xlsx").Sheets(1).Range("G6").Offset(x)
            Exit For
            
        Else: End If
    Next x
                
Workbooks("Job Labor Totals.xlsx").Saved = True
Workbooks("Job Labor Totals.xlsx").Close

'----------------------------------
'Stage 3: Get invoicing totals
'-----------------------------------
locatePath = targetPath & "\Backup Reports\Over-Under Billings.xlsx"
Application.ScreenUpdating = False

'Workbooks.Open (locatepath)
'Workbooks("Over-Under Billings.xlsx").Close

End Sub
