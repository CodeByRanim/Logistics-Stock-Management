# VBA Excel - Optimisation de la gestion des stocks et des prévisions de la demande

Ce repository présente une solution VBA pour automatiser la gestion des stocks et les prévisions de la demande dans le cadre de la chaîne logistique. Le fichier permet de calculer le stock de sécurité, les quantités de commande nécessaires et d'optimiser les prévisions en fonction des tendances passées.

## Fonctionnalités :
- Calcul automatique des niveaux de stock de sécurité.
- Prévisions des quantités à commander en fonction des ventes passées.
- Suivi des niveaux de stock et des dates de réapprovisionnement.
- Mise à jour automatique des stocks après réception des commandes.

## Installation :

1. Ouvrez votre fichier Excel.
2. Allez dans l'éditeur VBA (ALT + F11).
3. Importez le fichier `stock_management.bas` via `Fichier > Importer un fichier...`.
4. Insérez vos données de ventes et de stocks dans les feuilles `SalesData` et `StockData`.

## Utilisation :

1. Saisissez les données de ventes mensuelles dans la feuille `SalesData` (colonne A : date, colonne B : quantités vendues).
2. Saisissez les données de stocks actuels dans la feuille `StockData` (colonne A : produit, colonne B : niveau de stock).
3. Exécutez la macro `OptimizeStockManagement` pour obtenir des recommandations sur les niveaux de stock et les quantités à commander.

## Exemples de personnalisation :
- Personnaliser les critères de réapprovisionnement en fonction de la saisonnalité.
- Ajouter des critères pour gérer les délais de livraison des fournisseurs.

## Contribuer :

Les contributions sont les bienvenues ! Si vous avez des suggestions pour améliorer ce projet, n'hésitez pas à soumettre des pull requests ou à ouvrir des issues.
