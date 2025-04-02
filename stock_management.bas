' Module pour la gestion des stocks et des prévisions de demande dans la chaîne logistique
' Ce module contient des fonctions pour calculer le stock de sécurité, les prévisions de commandes et les niveaux de stock.

Sub OptimizeStockManagement()

    Dim wsSales As Worksheet
    Dim wsStock As Worksheet
    Dim lastRowSales As Long
    Dim lastRowStock As Long
    Dim i As Long
    Dim product As String
    Dim salesData As Variant
    Dim stockData As Variant
    Dim forecastDemand As Double
    Dim safetyStock As Double
    Dim orderQuantity As Double
    
    ' Définir les feuilles
    Set wsSales = ThisWorkbook.Sheets("SalesData")
    Set wsStock = ThisWorkbook.Sheets("StockData")
    
    ' Récupérer les données de ventes et de stocks
    lastRowSales = wsSales.Cells(wsSales.Rows.Count, 1).End(xlUp).Row
    salesData = wsSales.Range("A2:B" & lastRowSales).Value
    
    lastRowStock = wsStock.Cells(wsStock.Rows.Count, 1).End(xlUp).Row
    stockData = wsStock.Range("A2:B" & lastRowStock).Value
    
    ' Appliquer les calculs pour chaque produit
    For i = 1 To lastRowStock - 1
        product = stockData(i, 1)
        
        ' Prévision de la demande : moyenne des ventes des 3 derniers mois
        forecastDemand = CalculateForecast(product, salesData)
        
        ' Calcul du stock de sécurité : basé sur la demande moyenne et le délai de réapprovisionnement
        safetyStock = forecastDemand * 0.1 ' Exemple : 10% de la demande moyenne comme stock de sécurité
        
        ' Calcul de la quantité à commander : prévision de la demande - stock actuel + stock de sécurité
        orderQuantity = forecastDemand - stockData(i, 2) + safetyStock
        
        ' Si la quantité à commander est positive, afficher la quantité à commander
        If orderQuantity > 0 Then
            wsStock.Cells(i + 1, 3).Value = orderQuantity ' Affichage de la quantité à commander dans la colonne C
        Else
            wsStock.Cells(i + 1, 3).Value = 0 ' Aucun réapprovisionnement nécessaire
        End If
    Next i
    
    MsgBox "La gestion des stocks a été optimisée avec succès.", vbInformation, "Optimisation terminée"
    
End Sub

' Fonction pour calculer la prévision de la demande pour un produit donné
Function CalculateForecast(product As String, salesData As Variant) As Double
    Dim totalSales As Double
    Dim count As Long
    Dim i As Long
    
    totalSales = 0
    count = 0
    
    ' Parcourir les données de ventes pour le produit
    For i = 1 To UBound(salesData, 1)
        If salesData(i, 1) = product Then
            totalSales = totalSales + salesData(i, 2)
            count = count + 1
        End If
    Next i
    
    ' Calculer la prévision de la demande : moyenne des ventes
    If count > 0 Then
        CalculateForecast = totalSales / count
    Else
        CalculateForecast = 0 ' Si aucune donnée de vente n'est disponible, retour à zéro
    End If
End Function
