# expirymacro
Compares expiring items with data in zhlxplan
Sub expirymacro()
Dim expiry As Worksheet
Dim hrpr As Worksheet
Dim myunit As String

'must include tab with zhlxplan output name "hrpr"


Set expiry = Application.ActiveWorkbook.Sheets("expiry")
Set hrpr = Application.ActiveWorkbook.Sheets("hrpr")

lastrow = expiry.Cells(Rows.Count, 1).End(xlUp).Row
lrhrpr = hrpr.Cells(Rows.Count, 1).End(xlUp).Row


                        expiry.Cells(4, 12) = "Target Stock"
                        expiry.Cells(4, 13) = "ROP"
                        expiry.Cells(4, 14) = "0 - 3 Mo"
                        expiry.Cells(4, 15) = "3 - 6 Mo"
                        expiry.Cells(4, 16) = "6 - 12 Mo"
                         expiry.Cells(4, 11) = "Item"
                        expiry.Cells(4, 10) = "Unit"
                        expiry.Cells(5, 7) = "New Target"

For x = 1 To lastrow
unittest = InStr(expiry.Cells(x, 2), "1000") 'does cell have 10000
    If unittest >= 1 Then
        myunit = expiry.Cells(x, 2)  'if yes, than that is a unit
        For a = 1 To 5    'searching down from unit look for items
        theitem = expiry.Cells(x, 2).Offset(a, 0)
            For y = 1 To lrhrpr
                If myunit = hrpr.Cells(y, 2) And theitem = hrpr.Cells(y, 4) Then   'try to match unit and item
                        expiry.Cells(x, 2).Offset(a, 10) = hrpr.Cells(y, 7) 'target stock
                        expiry.Cells(x, 2).Offset(a, 11) = hrpr.Cells(y, 8) ' rop
                        expiry.Cells(x, 2).Offset(a, 12) = hrpr.Cells(y, 13) ' 0 to 3 months
                        expiry.Cells(x, 2).Offset(a, 13) = hrpr.Cells(y, 14) ' 3 to 6 months
                        expiry.Cells(x, 2).Offset(a, 14) = hrpr.Cells(y, 15) ' 6 to 12 months
                         expiry.Cells(x, 2).Offset(a, 9) = hrpr.Cells(y, 4)
                        expiry.Cells(x, 2).Offset(a, 8) = hrpr.Cells(y, 2)
                        
                End If
            Next y
        Next a
    
     
    
    
    End If

'
Next x
'''to confirm changes were made. Re-run HRPR and compare new target stock (Column H) and Target Stock from HRPR (Column L)
  If WorksheetFunction.CountA(Range("G9:G25")) = 0 Then
        
    Else
        For l = 6 To lastrow
        If expiry.Cells(l, 12) <> expiry.Cells(l, 7) Then
        expiry.Cells(l, 17) = "error"
        End If
        Next l
    End If
    




End Sub

