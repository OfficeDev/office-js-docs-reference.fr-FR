| Classe | Champs | Description |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|Adresse de la cellule qui contient la formule modifiée.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Représente la formule précédente, avant qu’elle n’a été modifiée.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positiontype)|Position d’insertion, dans le livre de calcul actuel, des nouvelles feuilles de calcul.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#relativeto)|Feuille de calcul du manuel actuel référencé pour le `WorksheetPositionType` paramètre.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetnamestoinsert)|Noms des feuilles de calcul individuelles à insérer.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|Description de texte de alt du tableau croisé dynamique.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|Titre de texte de alt du tableau croisé dynamique.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Définit si une ligne vide doit être affichée après chaque élément.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Texte qui est automatiquement rempli dans n’importe quelle cellule vide du tableau croisé dynamique si `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Spécifie si les cellules vides du tableau croisé dynamique doivent être remplies avec le `emptyCellText` .|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Définit le paramètre « Répéter toutes les étiquettes d’éléments » sur tous les champs du tableau croisé dynamique.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Spécifie si le tableau croisé dynamique affiche les en-têtes de champ (légendes de champ et les drop-downs de filtre).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Spécifie si le tableau croisé dynamique est actualisé à l’ouverture du manuel.|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents()](/javascript/api/excel/excel.range#getdirectdependents--)|Renvoie un objet qui représente la plage contenant tous les dépendants directs d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
||[getExtendedRange(direction: Excel. KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|Renvoie un objet de plage qui inclut la plage actuelle et jusqu’au bord de la plage, en fonction de la direction fournie.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#getmergedareasornullobject--)|Renvoie un objet RangeAreas qui représente les zones fusionnées dans cette plage.|
||[getRangeEdge(direction: Excel. KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|Renvoie un objet de plage qui est la cellule edge de la zone de données qui correspond à la direction fournie.|
|[Table](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|Resize the table to the new range.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel. InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#insertworksheetsfrombase64-base64file--options-)|Insère les feuilles de calcul spécifiées à partir d’un workbook source dans le workbook actuel.|
||[onActivated](/javascript/api/excel/excel.workbook#onactivated)|Se produit lorsque le workbook est activé.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#type)|Obtient le type de l’événement.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Se produit lorsqu’une ou plusieurs formules sont modifiées dans cette feuille de calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Se produit lorsqu’une ou plusieurs formules sont modifiées dans une feuille de calcul de cette collection.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Obtient un tableau `FormulaChangedEventDetail` d’objets, qui contient les détails sur toutes les formules modifiées.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle la formule a été modifiée.|
