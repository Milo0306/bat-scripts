Attribute VB_Name = "Module1"
Sub HighlightDuplicatesKeepMaxHH()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim i As Long
    Dim colDup As Integer, colHH As Integer
    Dim tempDict As Object
    
    ' Ustawienie aktywnego arkusza
    Set ws = ActiveSheet
    
    ' Znalezienie kolumn na podstawie nazw
    colDup = Application.Match("DUPLIKATY", ws.Rows(1), 0)
    colHH = Application.Match("HH", ws.Rows(1), 0)
    
    ' Znalezienie ostatniego wiersza
    lastRow = ws.Cells(ws.Rows.Count, colDup).End(xlUp).Row
    
    ' Je�li nie ma danych, zako�cz
    If lastRow < 2 Then
        MsgBox "Brak danych do przetworzenia.", vbExclamation
        Exit Sub
    End If
    
    ' Tworzymy s�ownik do przechowywania maksymalnych warto�ci HH dla ka�dej warto�ci DUPLIKATY
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Przeszukiwanie danych w celu znalezienia najwi�kszego HH dla ka�dej warto�ci DUPLIKATY
    For i = 2 To lastRow
        If ws.Cells(i, colDup).Value <> "" Then
            ' Konwersja warto�ci HH na liczb� (na wypadek, gdyby by�y traktowane jako tekst)
            ws.Cells(i, colHH).Value = Val(ws.Cells(i, colHH).Value)
            
            If Not dict.exists(ws.Cells(i, colDup).Value) Then
                dict.Add ws.Cells(i, colDup).Value, ws.Cells(i, colHH).Value
            Else
                ' Je�li HH jest wi�ksze, aktualizujemy warto��
                If ws.Cells(i, colHH).Value > dict(ws.Cells(i, colDup).Value) Then
                    dict(ws.Cells(i, colDup).Value) = ws.Cells(i, colHH).Value
                End If
            End If
        End If
    Next i

    ' Drugi s�ownik dla unikalnych wpis�w
    Set tempDict = CreateObject("Scripting.Dictionary")

    ' Kolorowanie duplikat�w zamiast usuwania
    For i = 2 To lastRow
        If ws.Cells(i, colDup).Value <> "" Then
            ' Je�li HH nie jest najwi�ksze dla danej warto�ci DUPLIKATY � KOLORUJEMY
            If ws.Cells(i, colHH).Value <> dict(ws.Cells(i, colDup).Value) Then
                Debug.Print "Koloruj� wiersz: "; ws.Cells(i, colDup).Value, "HH:", ws.Cells(i, colHH).Value
                ws.Rows(i).Interior.Color = RGB(255, 255, 0) ' ��ty kolor
            ' Je�li HH jest najwi�ksze, ale wpis ju� wyst�puje, kolorujemy jeden z duplikat�w
            ElseIf tempDict.exists(ws.Cells(i, colDup).Value) Then
                Debug.Print "Koloruj� powtarzaj�cy si� wiersz: "; ws.Cells(i, colDup).Value, "HH:", ws.Cells(i, colHH).Value
                ws.Rows(i).Interior.Color = RGB(255, 255, 0) ' ��ty kolor
            Else
                ' Zapisujemy unikalny wpis, �eby zostawi� tylko jeden
                tempDict.Add ws.Cells(i, colDup).Value, ws.Cells(i, colHH).Value
            End If
        End If
    Next i
    
    MsgBox "Duplikaty oznaczone kolorem ��tym.", vbInformation
End Sub

