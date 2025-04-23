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
    
    ' Jeœli nie ma danych, zakoñcz
    If lastRow < 2 Then
        MsgBox "Brak danych do przetworzenia.", vbExclamation
        Exit Sub
    End If
    
    ' Tworzymy s³ownik do przechowywania maksymalnych wartoœci HH dla ka¿dej wartoœci DUPLIKATY
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Przeszukiwanie danych w celu znalezienia najwiêkszego HH dla ka¿dej wartoœci DUPLIKATY
    For i = 2 To lastRow
        If ws.Cells(i, colDup).Value <> "" Then
            ' Konwersja wartoœci HH na liczbê (na wypadek, gdyby by³y traktowane jako tekst)
            ws.Cells(i, colHH).Value = Val(ws.Cells(i, colHH).Value)
            
            If Not dict.exists(ws.Cells(i, colDup).Value) Then
                dict.Add ws.Cells(i, colDup).Value, ws.Cells(i, colHH).Value
            Else
                ' Jeœli HH jest wiêksze, aktualizujemy wartoœæ
                If ws.Cells(i, colHH).Value > dict(ws.Cells(i, colDup).Value) Then
                    dict(ws.Cells(i, colDup).Value) = ws.Cells(i, colHH).Value
                End If
            End If
        End If
    Next i

    ' Drugi s³ownik dla unikalnych wpisów
    Set tempDict = CreateObject("Scripting.Dictionary")

    ' Kolorowanie duplikatów zamiast usuwania
    For i = 2 To lastRow
        If ws.Cells(i, colDup).Value <> "" Then
            ' Jeœli HH nie jest najwiêksze dla danej wartoœci DUPLIKATY › KOLORUJEMY
            If ws.Cells(i, colHH).Value <> dict(ws.Cells(i, colDup).Value) Then
                Debug.Print "Kolorujê wiersz: "; ws.Cells(i, colDup).Value, "HH:", ws.Cells(i, colHH).Value
                ws.Rows(i).Interior.Color = RGB(255, 255, 0) ' ¯ó³ty kolor
            ' Jeœli HH jest najwiêksze, ale wpis ju¿ wystêpuje, kolorujemy jeden z duplikatów
            ElseIf tempDict.exists(ws.Cells(i, colDup).Value) Then
                Debug.Print "Kolorujê powtarzaj¹cy siê wiersz: "; ws.Cells(i, colDup).Value, "HH:", ws.Cells(i, colHH).Value
                ws.Rows(i).Interior.Color = RGB(255, 255, 0) ' ¯ó³ty kolor
            Else
                ' Zapisujemy unikalny wpis, ¿eby zostawiæ tylko jeden
                tempDict.Add ws.Cells(i, colDup).Value, ws.Cells(i, colHH).Value
            End If
        End If
    Next i
    
    MsgBox "Duplikaty oznaczone kolorem ¿ó³tym.", vbInformation
End Sub

