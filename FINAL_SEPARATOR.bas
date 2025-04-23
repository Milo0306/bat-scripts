Attribute VB_Name = "FINAL_SEPARATOR"
Sub SplitByIDPHAcrossSheetsWithFolderSelection()
    Dim ws As Worksheet
    Dim uniqueIDs As Collection
    Dim employeeIDs
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
    Dim dataFound As Boolean
    Dim employeeSheetName As String
    Dim employeeSheet As Worksheet
    Dim temp As String
    
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
    Set employeeIDs = CreateObject("Scripting.Dictionary")
    employeeSheetName = "LISTA PH"
    Set employeeSheet = Sheets(employeeSheetName)
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = employeeSheetName Then
            ' Znajdowanie ostatniego wiersza w kolumnie A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column ' Znalazienie ostatniej kolumny
            
            ' Pobieranie unikalnych ID_PH w danym arkuszu
            On Error Resume Next
            For i = 2 To lastRow ' Zak³adaj¹c, ¿e dane zaczynaj¹ siê od drugiego wiersza
                temp = ws.Cells(i, 1).Value
                employeeIDs.Add ws.Cells(i, 1).Value, ws.Cells(i, 2).Value
            Next i
            On Error GoTo 0
        End If
    Next ws
    
    
    ' Iteracja po wszystkich zak³adkach w pliku, z pominiêciem arkusza MENU
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "MENU" And ws.Name <> employeeSheetName Then ' Sprawdzamy, czy arkusz nie jest MENU
            ' Znajdowanie ostatniego wiersza w kolumnie A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column ' Znalazienie ostatniej kolumny
            
            ' Pobieranie unikalnych ID_PH w danym arkuszu
            On Error Resume Next
            For i = 2 To lastRow ' Zak³adaj¹c, ¿e dane zaczynaj¹ siê od drugiego wiersza
                uniqueIDs.Add ws.Cells(i, 1).Value, CStr(ws.Cells(i, 1).Value)
            Next i
            On Error GoTo 0
        End If
    Next ws
    
    ' Iteracja przez unikalne ID_PH
    For Each ID In uniqueIDs
        ' Tworzenie nowego skoroszytu
        Set newWorkbook = Workbooks.Add
        
        ' Iteracja po arkuszach w pliku g³ównym, z pominiêciem arkusza MENU
        For Each ws In ThisWorkbook.Sheets
            If ws.Name <> "MENU" And ws.Name <> employeeSheetName Then ' Sprawdzamy, czy arkusz nie jest MENU
                ' Sprawdzenie, czy w arkuszu znajduj¹ siê dane dla danego ID_PH
                dataFound = False
                
                ' Tworzenie nowego arkusza w nowym pliku tylko wtedy, gdy dane dla ID_PH s¹ obecne
                Set newSheet = Nothing
                For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                    If ws.Cells(i, 1).Value = ID Then
                        If newSheet Is Nothing Then
                            ' Tworzenie nowego arkusza, jeœli jeszcze nie zosta³ stworzony
                            Set newSheet = newWorkbook.Sheets.Add
                            newSheet.Name = ws.Name ' Zachowanie nazwy arkusza
                            ws.Rows(1).Copy Destination:=newSheet.Rows(1) ' Kopiowanie nag³ówków
                        End If
                        ' Kopiowanie danych
                        newSheet.Rows(newSheet.Cells(newSheet.Rows.Count, "A").End(xlUp).Row + 1).Value = ws.Rows(i).Value
                        dataFound = True
                    End If
                Next i
                
                ' Jeœli nie znaleziono danych, nie dodawaj arkusza
                If Not dataFound And Not newSheet Is Nothing Then
                    newSheet.Delete
                End If
            End If
        Next ws
        
        ' Ustalanie nazwy pliku (ID_PH) i zapisanie nowego pliku
        If CStr(employeeIDs.Item(ID)) = "" Then
            outputFile = folderPath & "\" & CStr(ID) & ".xlsx" ' Zapisz plik w wybranym folderze
        Else
            temp = CStr(employeeIDs.Item(ID))
            outputFile = folderPath & "\" & CStr(employeeIDs.Item(ID)) & ".xlsx" ' Zapisz plik w wybranym folderze
            End If
        newWorkbook.SaveAs outputFile
        newWorkbook.Close False
        
    Next ID
    
    MsgBox "Podzia³ pliku zakoñczony!"
End Sub

