Attribute VB_Name = "IMPORTER_SEGREGATOR"
Sub ImportFilesToSheets()
    Dim folderPath As String
    Dim FileName As String
    Dim sheetName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim Prefix As String
    Dim OldNames As Variant
    Dim NewNames As Variant
    Dim i As Long
    Dim Col As Long
    Dim Dict As Object
    Dim ID_PH_Col As Long

    ' Prefiks, kt�ry ma by� pomijany
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

    ' Wyszukiwanie plik�w w folderze
    FileName = Dir(folderPath & "*.xlsx") ' Mo�na doda� inne rozszerzenia jak .csv, .xls, je�li potrzebne
    
    ' Definiowanie starego i nowego nazewnictwa kolumn
    OldNames = Array("Kod sprzeda�owy", "id_przedstawiciel", "Kod APS")  ' Stare nazwy kolumn
    NewNames = Array("ID_PH", "ID_PH", "ID_PH")  ' Nowe nazwy kolumn
    
    ' Tworzymy s�ownik do mapowania nazw kolumn
    Set Dict = CreateObject("Scripting.Dictionary")
    For i = LBound(OldNames) To UBound(OldNames)
        Dict.Add OldNames(i), NewNames(i)
    Next i
    
    ' P�tla przez wszystkie pliki w folderze
    Do While FileName <> ""
        ' Sprawdzanie, czy plik jest odpowiedni
        If FileName <> "" Then
            ' Usuni�cie prefiksu z nazwy pliku, je�li jest obecny
            If Left(FileName, Len(Prefix)) = Prefix Then
                sheetName = Mid(FileName, Len(Prefix) + 1, Len(FileName) - Len(Prefix) - 5) ' Usuwa prefiks D5224_ i rozszerzenie .xlsx
            Else
                sheetName = Left(FileName, Len(FileName) - 5) ' Usuwa rozszerzenie .xlsx
            End If
            
            ' Tworzenie nowego arkusza z nazw� pliku bez prefiksu i rozszerzenia
            Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
            ws.Name = sheetName
            
            ' Otw�rz plik i skopiuj zawarto��
            With Workbooks.Open(folderPath & FileName)
                ' Zak�ada, �e dane s� w pierwszym arkuszu
                .Sheets(1).UsedRange.Copy Destination:=ws.Range("A1")
                
                ' Zamiana nazw kolumn
                lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                For i = 1 To ws.UsedRange.Columns.Count
                    ' Sprawdzenie, czy nag��wek kolumny znajduje si� w s�owniku
                    If Dict.Exists(ws.Cells(1, i).Value) Then
                        ws.Cells(1, i).Value = Dict(ws.Cells(1, i).Value) ' Zmieniamy nazw� kolumny
                    End If
                Next i

                ' Szukamy kolumny "ID_PH" i przesuwamy j� na pierwsz� pozycj�
                For i = 1 To ws.UsedRange.Columns.Count
                    If ws.Cells(1, i).Value = "ID_PH" Then
                        ID_PH_Col = i
                        ' Sprawdzamy, czy kolumna "ID_PH" nie jest ju� pierwsza
                        If ID_PH_Col > 1 Then
                            ' Przenosimy ca�� kolumn� o jedno miejsce w lewo
                            ws.Columns(ID_PH_Col).Cut
                            ws.Columns(1).Insert Shift:=xlToRight
                        End If
                        Exit For
                    End If
                Next i
                
                .Close False
            End With
        End If
        FileName = Dir ' Kolejny plik
    Loop
    
    MsgBox "Pliki zosta�y zaimportowane, nag��wki zosta�y zamienione, a kolumna 'ID_PH' zosta�a przeniesiona na pierwsz� pozycj�!"
End Sub

