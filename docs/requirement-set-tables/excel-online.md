| Classe | Champs | Description |
|:---|:---|:---|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Active cette vue de feuille.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Supprime l’affichage Feuille de la feuille de calcul.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Crée une copie de cette vue de feuille.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtient ou définit le nom de l’affichage Feuille.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Crée un affichage feuille avec le nom donné.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Crée et active un nouvel affichage de feuille temporaire.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Quitte l’affichage feuille actif.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Obtient la vue de feuille de calcul active.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Obtient le nombre d’affichages de feuille dans cette feuille de calcul.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Obtient une vue de feuille à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Obtient une vue de feuille par son index dans la collection.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Range](/javascript/api/excel/excel.range)|[getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|Renvoie un objet de plage qui inclut la plage actuelle et jusqu’au bord de la plage, en fonction de la direction fournie.|
||[getMergedAreas()](/javascript/api/excel/excel.range#getmergedareas--)|Renvoie un `RangeAreas` objet qui représente les zones fusionnées dans cette plage.|
||[getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|Renvoie un objet de plage qui est la cellule edge de la zone de données qui correspond à la direction fournie.|
|[Tableau](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|Resize the table to the new range.|
|[Feuille de calcul](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Renvoie une collection d’affichages de feuille présents dans la feuille de calcul.|
