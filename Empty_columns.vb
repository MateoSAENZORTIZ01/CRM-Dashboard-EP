Sub GenererTableauVidesDansNouvelleFeuille()
    Dim ws As Worksheet, wsResult As Worksheet
    Dim lastCol As Long, lastRow As Long
    Dim resultRow As Long
    Dim pipelineDict As Object
    Dim pipelineCol As Range, cell As Range
    Dim colIndex As Long, emptyCount As Long, totalLignes As Long
    Dim pipelineValue As String
    
    ' Définir la feuille active (données sources)
    Set ws = ActiveSheet

    ' Vérifier si la feuille "Analyse Cellules Vides" existe déjà et la supprimer
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Analyse Cellules Vides").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Créer une nouvelle feuille pour les résultats
    Set wsResult = ThisWorkbook.Sheets.Add
    wsResult.Name = "Analyse Cellules Vides"

    ' Trouver la dernière colonne et la dernière ligne dans la feuille source
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Créer un dictionnaire pour stocker les valeurs uniques de la colonne Pipeline
    Set pipelineDict = CreateObject("Scripting.Dictionary")

    ' Parcourir la colonne Pipeline pour récupérer les valeurs uniques
    Set pipelineCol = ws.Range("B2:B" & lastRow)
    For Each cell In pipelineCol
        If Not pipelineDict.exists(cell.Value) And cell.Value <> "" Then
            pipelineDict.Add cell.Value, Nothing
        End If
    Next cell

    ' Écrire les en-têtes dans la feuille "Analyse Cellules Vides"
    wsResult.Cells(1, 1).Value = "Pipeline"
    wsResult.Cells(1, 2).Value = "Colonne"
    wsResult.Cells(1, 3).Value = "Cellules Vides"
    wsResult.Cells(1, 4).Value = "Total Lignes"

    ' Appliquer un format en gras aux titres
    wsResult.Range("A1:D1").Font.Bold = True

    ' Initialiser la ligne de sortie
    resultRow = 2

    ' Boucler sur chaque valeur de Pipeline
    Dim key As Variant
    For Each key In pipelineDict.keys
        pipelineValue = key
        
        ' Calculer le total de lignes pour ce pipeline
        totalLignes = Application.WorksheetFunction.CountIf(ws.Range("B2:B" & lastRow), pipelineValue)

        ' Boucler sur toutes les colonnes à partir de la deuxième
        For colIndex = 2 To lastCol
            ' Compter les cellules vides pour ce pipeline dans la colonne courante
            emptyCount = Application.WorksheetFunction.CountIfs(ws.Range("B2:B" & lastRow), pipelineValue, ws.Range(ws.Cells(2, colIndex), ws.Cells(lastRow, colIndex)), "")

            ' Ajouter au tableau des résultats dans la nouvelle feuille
            wsResult.Cells(resultRow, 1).Value = pipelineValue
            wsResult.Cells(resultRow, 2).Value = ws.Cells(1, colIndex).Value ' Nom de la colonne
            wsResult.Cells(resultRow, 3).Value = emptyCount ' Nombre de cellules vides
            wsResult.Cells(resultRow, 4).Value = totalLignes ' Nombre total de lignes pour ce pipeline
            resultRow = resultRow + 1
        Next colIndex
    Next key

    ' Ajuster la largeur des colonnes pour une meilleure lisibilité
    wsResult.Columns("A:D").AutoFit

    MsgBox "Tableau généré dans la feuille 'Analyse Cellules Vides' avec succès !", vbInformation
End Sub
