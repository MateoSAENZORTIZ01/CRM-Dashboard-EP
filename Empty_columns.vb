Sub GenererTableauVidesTousPipelinesComplet()
    Dim ws As Worksheet
    Dim lastCol As Long, lastRow As Long
    Dim resultRow As Long
    Dim pipelineDict As Object
    Dim pipelineCol As Range, cell As Range
    Dim colIndex As Long, emptyCount As Long
    Dim pipelineValue As String
    
    ' Définir la feuille active
    Set ws = ActiveSheet

    ' Trouver la dernière colonne et la dernière ligne
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

    ' Déterminer où commencer le tableau des résultats
    resultRow = lastRow + 2
    ws.Cells(resultRow, 1).Value = "Pipeline"
    ws.Cells(resultRow, 2).Value = "Colonne"
    ws.Cells(resultRow, 3).Value = "Cellules Vides"

    ' Appliquer un format en gras aux titres
    ws.Range(ws.Cells(resultRow, 1), ws.Cells(resultRow, 3)).Font.Bold = True

    ' Initialiser la ligne de sortie
    resultRow = resultRow + 1

    ' Boucler sur chaque valeur de Pipeline
    Dim key As Variant
    For Each key In pipelineDict.keys
        pipelineValue = key

        ' Boucler sur toutes les colonnes à partir de la deuxième
        For colIndex = 2 To lastCol
            ' Compter les cellules vides pour ce pipeline dans la colonne courante
            emptyCount = Application.WorksheetFunction.CountIfs(ws.Range("B2:B" & lastRow), pipelineValue, ws.Range(ws.Cells(2, colIndex), ws.Cells(lastRow, colIndex)), "")

            ' Ajouter au tableau des résultats (même si emptyCount = 0)
            ws.Cells(resultRow, 1).Value = pipelineValue
            ws.Cells(resultRow, 2).Value = ws.Cells(1, colIndex).Value ' Nom de la colonne
            ws.Cells(resultRow, 3).Value = emptyCount ' Nombre de cellules vides
            resultRow = resultRow + 1
        Next colIndex
    Next key

    MsgBox "Tableau généré avec succès !", vbInformation
End Sub

