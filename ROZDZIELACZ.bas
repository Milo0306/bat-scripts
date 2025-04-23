Attribute VB_Name = "ROZDZIELACZ"
Sub SplitByIDPHAcrossSheetsWithFolderSelection()
    Dim ws As Worksheet
    Dim uniqueIDs As Collection
    Dim ID As Variant
    Dim newWorkbook As Workbook
    Dim newSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentRow As Long
    Dim originalRange As Range
    Dim outputFile As String
    Dim sheetName As String
    Dim lastColumn As Long
    Dim folderPath As String
    Dim dialog As FileDialog
    
    ' Inicjalizacja okna dialogowego wyboru folderu
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    dialog.Title = "Wybierz folder do zapisania plików"
    
    ' Poka¿ okno dialogowe i zapisz wybran¹ œcie¿kê
    If dialog.Show = -1 Then
        folderPath = dialog.SelectedItems(1)
    Else
        MsgBox "Nie wybrano folderu. Proces zostanie przerwany."
        Exit Sub
    End If
    
    ' Kolekcja do przechowywania unikalnych ID_PH
    Set uniqueIDs = New Collection
    
    ' Iteracja po wszystkich zak³adkach w pliku
    For Each ws In ThisWorkbook.Sheets
        ' Znajdowanie ostatniego wiersza w kolumnie A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column ' Znalazienie ostatniej kolumny
        
        ' Pobieranie unikalnych ID_PH w danym arkuszu
        On Error Resume Next
        For i = 2 To lastRow ' Zak³adaj¹c, ¿e dane zaczynaj¹ siê od drugiego wiersza
            uniqueIDs.Add ws.Cells(i, 1).Value, CStr(ws.Cells(i, 1).Value)
        Next i
        On Error GoTo 0
    Next ws
    
    ' Iteracja przez unikalne ID_PH
    For Each ID In uniqueIDs
        ' Tworzenie nowego skoroszytu
        Set newWorkbook = Workbooks.Add
        
        ' Iteracja po arkuszach w pliku g³ównym
        For Each ws In ThisWorkbook.Sheets
            ' Tworzenie nowego arkusza w nowym pliku
            Set newSheet = newWorkbook.Sheets.Add
            newSheet.Name = ws.Name ' Zachowanie nazwy arkusza

            ' Kopiowanie nag³ówków
            ws.Rows(1).Copy Destination:=newSheet.Rows(1)
            
            ' Wyszukiwanie danych dla danego ID_PH i kopiowanie
            currentRow = 2
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                If ws.Cells(i, 1).Value = ID Then
                    ws.Rows(i).Copy Destination:=newSheet.Rows(currentRow)
                    currentRow = currentRow + 1
                End If
            Next i
        Next ws
        
        ' Ustalanie nazwy pliku (ID_PH) i zapisanie nowego pliku
        outputFile = folderPath & "\" & CStr(ID) & ".xlsx" ' Zapisz plik w wybranym folderze
        newWorkbook.SaveAs outputFile
        newWorkbook.Close False
    Next ID
    
    MsgBox "Podzia³ pliku zakoñczony!"
End Sub

