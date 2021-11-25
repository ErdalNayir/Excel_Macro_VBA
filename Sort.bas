Sub KesinSonuc()
'
' KesinSonuc Macro
'
' Keyboard Shortcut: Ctrl+Shift+S
'
   Dim ilk As Worksheet
   Dim sonra As Worksheet
   
   Dim count_row As Integer
   Dim sayac As Integer
   Dim basaDon As Integer
   

   Set ilk = ThisWorkbook.Sheets(1)
   Set sonra = ThisWorkbook.Sheets(2)
   
   sayac = 1
   basaDon = 1
    
   ilk.Activate
   
   count_row = WorksheetFunction.CountA(Range("A1", Range("A1").End(xlDown)))
    
   For i = 1 To count_row
   
        sonra.Cells(sayac, basaDon) = ilk.Cells(i, 1).Text
        
        basaDon = basaDon + 1
        
        If i Mod 7 = 0 Then
            sayac = sayac + 1
            basaDon = 1
            
        End If
        
   Next i
   
sonra.Activate
      
End Sub
