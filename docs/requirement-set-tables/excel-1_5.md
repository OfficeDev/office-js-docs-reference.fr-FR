| Classe | Champs | Description |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete--)|Supprime la partie XML personnalisée.|
||[getXml()](/javascript/api/excel/excel.customxmlpart#getxml--)|Obtient l’intégralité du contenu XML de la partie XML personnalisée.|
||[id](/javascript/api/excel/excel.customxmlpart#id)|ID de la partie XML personnalisée.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceuri)|URI d’espace de noms de la partie XML personnalisée.|
||[setXml(xml: string)](/javascript/api/excel/excel.customxmlpart#setxml-xml-)|Définit l’intégralité du contenu XML de la partie XML personnalisée.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add(xml: string)](/javascript/api/excel/excel.customxmlpartcollection#add-xml-)|Ajoute une nouvelle partie XML personnalisée au classeur.|
||[getByNamespace(namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#getbynamespace-namespaceuri-)|Obtient une nouvelle collection limitée de parties XML personnalisées dont les espaces de noms correspondent à l’espace de noms donné.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getcount--)|Obtient le nombre de parties XML personnalisées dans la collection.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitem-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitemornullobject-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getcount--)|Obtient le nombre de parties CustomXML dans cette collection.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitem-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitemornullobject-id-)|Obtient une partie XML personnalisée en fonction de son ID.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitem--)|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitemornullobject--)|Si la collection contient exactement un élément, cette méthode le renvoie.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|ID du tableau croisé dynamique.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#runtime)|[Ensemble d’api : ExcelApi 1.5]|
|[Runtime](/javascript/api/excel/excel.runtime)||[Classeur](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customxmlparts)|Représente la collection de parties XML personnalisées contenues dans ce manuel.|
|[Feuille de calcul](/javascript/api/excel/excel.worksheet)|[getNext(visibleOnly?: booléen)](/javascript/api/excel/excel.worksheet#getnext-visibleonly-)|Obtient la feuille de calcul qui suit celle-ci.|
||[getNextOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getnextornullobject-visibleonly-)|Obtient la feuille de calcul qui suit celle-ci.|
||[getPrevious(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getprevious-visibleonly-)|Obtient la feuille de calcul qui précède celle-ci.|
||[getPreviousOrNullObject(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getpreviousornullobject-visibleonly-)|Obtient la feuille de calcul qui précède celle-ci.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst(visibleOnly?: booléen)](/javascript/api/excel/excel.worksheetcollection#getfirst-visibleonly-)|Obtient la première feuille de calcul dans la collection.|
||[getLast(visibleOnly?: booléen)](/javascript/api/excel/excel.worksheetcollection#getlast-visibleonly-)|Obtient la dernière feuille de calcul dans la collection.|
