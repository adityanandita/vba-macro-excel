Sub ExtractData()
    Dim wb As Workbook
    Dim sourceFile As String
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim lastRow As Long
    Dim l_row As Long
    Dim i As Long
    
    ' Mematikan semua notifikasi
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Setting target sheet
    Set targetSheet = ThisWorkbook.Sheets("Data")
      
    ' Memilih file excel yang ingin diambil datanya
    sourceFile = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", Title:="Pilih file SALR")

    If sourceFile = "False" Then
        Exit Sub
    End If
    
    ' Menghitung last row dari datasheet
    l_row = targetSheet.Cells(Rows.Count, "B").End(xlUp).Row
    
    ' Menghapus data yang sudah ada untuk diupdate
    If sourceFile = "True" Then
        targetSheet.Range(Cells(3, "A"), Cells(l_row, "L")).ClearContents
    End If
        
    ' Membuka file excel yang dipilih
    Set wb = Workbooks.Open(sourceFile)
    
    ' Setting sumber sheet
    Set sourceSheet = wb.Sheets(1)
    
    ' Menyalin data kolom dari sumber sheet ke target sheet
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "D").End(xlUp).Row
    sourceSheet.Range("D7:D" & lastRow).Copy targetSheet.Range("A3")
    sourceSheet.Range("Y7:Y" & lastRow).Copy targetSheet.Range("B3")
    sourceSheet.Range("F7:F" & lastRow).Copy targetSheet.Range("C3")
    sourceSheet.Range("L7:L" & lastRow).Copy targetSheet.Range("D3")
    sourceSheet.Range("Q7:Y" & lastRow).Copy targetSheet.Range("E3")
    sourceSheet.Range("T7:T" & lastRow).Copy targetSheet.Range("F3")
    sourceSheet.Range("W7:W" & lastRow).Copy targetSheet.Range("G3")
    sourceSheet.Range("Y7:Y" & lastRow).Copy targetSheet.Range("H3")
    sourceSheet.Range("Z7:Z" & lastRow).Copy targetSheet.Range("I3")
    sourceSheet.Range("AA7:AA" & lastRow).Copy targetSheet.Range("J3")
    sourceSheet.Range("AB7:AB" & lastRow).Copy targetSheet.Range("K3")
    sourceSheet.Range("AC7:AC" & lastRow).Copy targetSheet.Range("L3")
    
    ' Menghapus setiap baris yang kolom DocumentNo/CoCd-nya Kosong atau "DocumentNo" atau "CoCd"
    For i = lastRow To 3 Step -1
        If targetSheet.Range("B" & i).Value = "" Or targetSheet.Range("B" & i).Value = "DocumentNo" Or targetSheet.Range("A" & i).Value = "" Or targetSheet.Range("A" & i).Value = "CoCd" Then
            targetSheet.Rows(i).Delete
        End If
    Next i
    
    'Menjumlahkan kolom H dan kolom I
    row_last = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row
    targetSheet.Range("I1").Value = WorksheetFunction.Sum(targetSheet.Range("I3:I" & row_last))
    targetSheet.Range("J1").Value = WorksheetFunction.Sum(targetSheet.Range("J3:J" & row_last))
    
    'Membuat Cell H1 dan I1 menjadi Bold
    targetSheet.Range("I1").Font.Bold = True
    targetSheet.Range("J1").Font.Bold = True
    
    ' Menutup file sumber
    wb.Close SaveChanges = False
End Sub