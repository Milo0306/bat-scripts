Attribute VB_Name = "CLEAN_FILE"
Sub DeleteAllSheetsExceptMenuAndListaPH()
    Dim ws As Worksheet
    
    ' Iteracja po wszystkich arkuszach
    For Each ws In ThisWorkbook.Sheets
        ' Sprawdza, czy arkusz nie jest MENU ani Dashboard
        If ws.Name <> "MENU" And ws.Name <> "LISTA PH" Then
            Application.DisplayAlerts = False 
            ws.Delete
            Application.DisplayAlerts = True 
        End If
    Next ws
    
    MsgBox "Wszystkie arkusze poza MENU i LISTA PH zosta�y usuni�te."
End Sub


