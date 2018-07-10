Attribute VB_Name = "funzioni"
Function SelezionaFile()

    Dim fPath As Variant

    fPath = Application.GetOpenFilename(Title:="Selezionare il file da aprire")
    If fPath = False Then Exit Function
    
    SelezionaFile = fPath

End Function

Function NumFogli(dir)

    Dim nFogli As Integer
    
        'nFogli: numero fogli presenti nel file input
        nFogli = Sheets.Count
  
    NumFogli = nFogli

End Function

Function NomiFogli(dir, nFogli)

    ReDim nomeFogli(nFogli - 2) As String
       
    'creazione del vettore nomi dei fogli del file di input
    For i = 1 To nFogli - 2
        nome_i = Worksheets(i + 1).Name
        nomeFogli(i) = nome_i
    Next i
   
    NomiFogli = nomeFogli

End Function

Function leggiDati(dir, nFogli, nomeFogli)

    Dim nCol As Integer
    Dim i As Integer
       
        'nFogli: numero fogli presenti nel file input
        nFogli = Sheets.Count
        ReDim NomiFogli(nFogli - 1) As String
        
        'creazione del vettore contenente il numero di righe piene per ogni foglio
        ReDim nRigPiene(nFogli - 2)
        
        For i = 1 To nFogli - 2
            nRigPiene(i) = Worksheets(i + 1).Cells(Rows.Count, 11).End(xlUp).Row
        Next i
        
        'creazione della variabile Collection "dati"
        ReDim dati(nFogli - 2)
        
        For i = 1 To nFogli - 2
            dati(i) = Worksheets(i + 1).Range(Worksheets(i + 1).Cells(3, 11), Worksheets(i + 1).Cells(nRigPiene(i), 15))
        Next i
        
    
    leggiDati = dati

End Function

Sub ScriviDati(datiCantiere, nFogli, nomeFogli)

    'pulizia del foglio "Misure_Reali"
    Worksheets("Misure_Reali").Cells.Clear

    'formattazione colonne foglio
    Worksheets("Misure_Reali").Range(Columns(1), Columns(9 * nFogli)).ColumnWidth = 20
    Worksheets("Misure_Reali").Rows(1).RowHeight = 30
    
    'scrittura intestazione tabella
    j = 0
    For i = 1 To (UBound(nomeFogli))
        Worksheets("Misure_Reali").Cells(1, j + 1) = "Data"
        Worksheets("Misure_Reali").Cells(1, j + 2) = nomeFogli(i) & " " & "Coordinate_TPS E"
        Worksheets("Misure_Reali").Cells(1, j + 4) = "Data"
        Worksheets("Misure_Reali").Cells(1, j + 5) = nomeFogli(i) & " " & "Coordinate_TPS N"
        Worksheets("Misure_Reali").Cells(1, j + 7) = "Data"
        Worksheets("Misure_Reali").Cells(1, j + 8) = nomeFogli(i) & " " & "Coordinate_TPS H"
    
        j = j + 9
    Next i
    
    'scrittura dati tabella
    k = 0
    For i = 1 To UBound(datiCantiere)
        'formattazione colonne come numero a 5 cifre decimali
        Worksheets("Misure_Reali").Columns(k + 2).NumberFormat = "0.00000"
        Worksheets("Misure_Reali").Columns(k + 5).NumberFormat = "0.00000"
        Worksheets("Misure_Reali").Columns(k + 8).NumberFormat = "0.00000"
        For j = 1 To UBound(datiCantiere(i))
            Worksheets("Misure_Reali").Cells(j + 1, k + 1) = datiCantiere(i)(j, 1)
            Worksheets("Misure_Reali").Cells(j + 1, k + 2) = datiCantiere(i)(j, 3)
            Worksheets("Misure_Reali").Cells(j + 1, k + 4) = datiCantiere(i)(j, 1)
            Worksheets("Misure_Reali").Cells(j + 1, k + 5) = datiCantiere(i)(j, 4)
            Worksheets("Misure_Reali").Cells(j + 1, k + 7) = datiCantiere(i)(j, 1)
            Worksheets("Misure_Reali").Cells(j + 1, k + 8) = datiCantiere(i)(j, 5)
        Next j
    k = k + 9
    Next i

End Sub
