Option Explicit

Dim Target_Folder As String
Dim wsSource As Worksheet, wsHelper As Worksheet
Dim lastRow As Long, LastColumn As Long

Sub SplitDataset()
    
    Dim collectionUniqueList As Collection
    Dim i As Long
    
    Set collectionUniqueList = New Collection
    
    Set wsSource = ThisWorkbook.Worksheets("Data")
    Set wsHelper = ThisWorkbook.Worksheets("Helper")
    
    ' Mengambil nilai Target_Folder dari sel D4 di lembar kerja "Helper"
    Target_Folder = wsHelper.Range("D4").Value
    
    ' Clear Helper Worksheet
    wsHelper.Cells.ClearContents
    
    With wsSource
        .AutoFilterMode = False
        
        lastRow = .Cells(Rows.Count, "A").End(xlUp).Row
        LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
        
        If .Range("A2").Value = "" Then
            GoTo Cleanup
        End If
        
        Call Init_Unique_List_Collection(collectionUniqueList, lastRow)
        
        Application.DisplayAlerts = False
        
        For i = 1 To collectionUniqueList.Count
                SplitWorksheet (collectionUniqueList.Item(i))
        Next i
        
        ActiveSheet.AutoFilterMode = False
        
    End With

Cleanup:

    Application.DisplayAlerts = True
    Set collectionUniqueList = Nothing
    Set wsSource = Nothing
    Set wsHelper = Nothing

End Sub

Private Sub Init_Unique_List_Collection(ByRef col As Collection, ByVal SourceWS_LastRow As Long)
    
    Dim lastRow As Long, RowNumber As Long
    
    ' Unique List Column
    wsSource.Range("D2:D" & SourceWS_LastRow).Copy wsHelper.Range("A1")
    
    With wsHelper
        
        If Len(Trim(.Range("A1").Value)) > 0 Then
            
            lastRow = .Cells(Rows.Count, "A").End(xlUp).Row
            
            .Range("A1:A" & lastRow).RemoveDuplicates 1, xlNo
            
            lastRow = .Cells(Rows.Count, "A").End(xlUp).Row
            
            .Range("A1:A" & lastRow).Sort.Range("A1"), Header:=xlNo
            
            lastRow = .Cells(Rows.Count, "A").End(xlUp).Row
            
            On Error Resume Next
            For RowNumber = 1 To lastRow
                col.Add .Cells(RowNumber, "A").Value, CStr(.Cells(RowNumber, "A").Value)
            Next RowNumber
           
        End If
    
    End With
    
End Sub

Private Sub SplitWorksheet(ByVal Category_Name As Variant)
    
    Dim wbTarget As Workbook
    Dim MappingSheet As Worksheet
    Dim LastRow_Mapping As Long
    Dim i As Long
    Dim File_Name As String
    
    Set MappingSheet = ThisWorkbook.Sheets("Mapping")
    
    LastRow_Mapping = MappingSheet.Cells(MappingSheet.Rows.Count, "A").End(xlUp).Row
    For i = 1 To LastRow_Mapping
        If MappingSheet.Cells(i, 1).Value = Category_Name Then
            File_Name = MappingSheet.Cells(i, 2).Value
            Exit For
        End If
    Next i
    
    If File_Name = "" Then
        File_Name = Category_Name
    End If
    
            
    Set wbTarget = Workbooks.Add
    
    With wsSource
        
        With .Range(.Cells(1, 1), .Cells(lastRow, LastColumn))
            .AutoFilter .Range("D1").Column, Category_Name
            
            .Copy
            
            'wbTarget.Worksheets(1).PasteSpecial xlValues
            wbTarget.Worksheets(1).Paste
            wbTarget.Worksheets(1).Name = Category_Name
            
            wbTarget.SaveAs Target_Folder & File_Name & ".xlsx", 51
            wbTarget.Close False
            
        End With
        
    End With
    
    Set wbTarget = Nothing
    wsSource.AutoFilterMode = False
    
End Sub
