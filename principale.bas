Attribute VB_Name = "principale"
Sub crea_file_input()
    
    ' dichiarazione variabili
    Dim nomeFileIn As String
    Dim dir As String
    Dim nFogli As Integer
    Dim fileIn As Workbook

    
    'creazione path file input
    dir = SelezionaFile()
    
    'apertura file dei dati di cantiere
    Set fileIn = Workbooks.Open(dir, True, True)
    
        'lettura numero di fogli presenti nel file dati di cantiere
        nFogli = NumFogli(dir)
    
        'lettura del nome dei fogli presenti nel file dati di cantiere
        ReDim nomeFogli(nFogli - 2) As String
        nomeFogli = NomiFogli(dir, nFogli)
    
        ' organizzazione dei dati provenienti dal cantiere in una variabile di tipo "Collection"
        datiCantiere = leggiDati(dir, nFogli, nomeFogli)
    
    'chiusura file dei dati di cantiere
    fileIn.Close
    
    'scrittura dati su file "input_elab_macro" foglio "Misure_Reali"
    Call ScriviDati(datiCantiere, nFogli, nomeFogli)
    
End Sub
