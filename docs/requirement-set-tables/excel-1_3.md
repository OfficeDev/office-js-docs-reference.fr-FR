| Classe | Champs | Description |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|Supprime la liaison.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add(range: Range \| string, bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Ajoute une nouvelle liaison à une plage spécifique.|
||[addFromNamedItem(name: string, bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Ajoute une nouvelle liaison basée sur un élément nommé dans le classeur.|
||[addFromSelection(bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Ajoute une nouvelle liaison basée sur la sélection en cours.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Nom du tableau croisé dynamique.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|Feuille de calcul contenant le tableau croisé dynamique.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|Actualise le tableau croisé dynamique.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|Obtient un tableau croisé dynamique par nom.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[refreshAll()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|Actualise tous les tableaux croisés dynamiques de la collection.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView()](/javascript/api/excel/excel.range#getvisibleview--)|Représente les lignes visibles de la plage en cours.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|Représente la formule dans le style de notation R1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|Obtient la plage parent associée à l’actuel `RangeView` .|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|Représente le code de format de nombre d’Excel pour une cellule donnée.|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|Représente les adresses de cellule du `RangeView` .|
||[columnCount](/javascript/api/excel/excel.rangeview#columncount)|Nombre de colonnes visibles.|
||[index](/javascript/api/excel/excel.rangeview#index)|Renvoie une valeur qui représente l’index du `RangeView` .|
||[rowCount](/javascript/api/excel/excel.rangeview#rowcount)|Nombre de lignes visibles.|
||[rows](/javascript/api/excel/excel.rangeview#rows)|Représente une collection d’affichages de plage associés à la plage.|
||[text](/javascript/api/excel/excel.rangeview#text)|Valeurs de texte de la plage spécifiée.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|Représente le type de données de chaque cellule.|
||[values](/javascript/api/excel/excel.rangeview#values)|Représente les valeurs brutes de l’affichage de plage spécifié.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|Obtient une `RangeView` ligne via son index.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Tableau](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|Spécifie si la première colonne contient une mise en forme spéciale.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|Spécifie si la dernière colonne contient une mise en forme spéciale.|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|Spécifie si les colonnes indiquent une mise en forme à bandes dans laquelle les colonnes impaires sont mises en surbrillante différemment des colonnes impaires, pour faciliter la lecture du tableau.|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|Spécifie si les lignes indiquent une mise en forme à bandes dans laquelle les lignes impaires sont mises en surbrillante différemment des lignes impaires, pour faciliter la lecture du tableau.|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|Spécifie si les boutons de filtre sont visibles en haut de chaque en-tête de colonne.|
|[Classeur](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivottables)|Représente une collection de tableaux croisés dynamiques associés au classeur.|
|[Feuille de calcul](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivottables)|Collection de tableaux croisés dynamiques qui font partie de la feuille de calcul.|
