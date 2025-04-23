Attribute VB_Name = "IMPORTER"
Sub ImportFilesToSheets()
    Dim folderPath As String
    Dim FileName As String
    Dim sheetName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim Prefix As String

    ' Prefiks, który ma byæ pomijany
    Prefix = "D5224_"
    
    ' Okno dialogowe do wyboru folderu
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Wybierz folder z plikami"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With
    
    ' Ustawienie nowego skoroszytu jako aktywnego
    Set wb = ThisWorkbook

    ' Wyszukiwanie plików w folderze
    FileName = Dir(folderPath & "*.xlsx") ' Mo¿na dodaæ inne rozszerzenia jak .csv, .xls, jeœli potrzebne
    
    ' Pêtla przez wszystkie pliki w folderze
    Do While FileName <> ""
        ' Sprawdzanie, czy plik jest odpowiedni
        If FileName <> "" Then
            ' Usuniêcie prefiksu z nazwy pliku, jeœli jest obecny
            If Left(FileName, Len(Prefix)) = Prefix Then
                sheetName = Mid(FileName, Len(Prefix) + 1, Len(FileName) - Len(Prefix) - 5) ' Usuwa prefiks D5224_ i rozszerzenie .xlsx
            Else
                sheetName = Left(FileName, Len(FileName) - 5) ' Usuwa rozszerzenie .xlsx
            End If
            
            ' Tworzenie nowego arkusza z nazw¹ pliku bez prefiksu i rozszerzenia
            Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
            ws.Name = sheetName
            
            ' Otwórz plik i skopiuj zawartoœæ
            With Workbooks.Open(folderPath & FileName)
                ' Zak³ada, ¿e dane s¹ w pierwszym arkuszu
                .Sheets(1).UsedRange.Copy Destination:=ws.Range("A1")
                .Close False
            End With
        End If
        FileName = Dir ' Kolejny plik
    Loop
    
    MsgBox "Pliki zosta³y zaimportowane!"
End Sub

