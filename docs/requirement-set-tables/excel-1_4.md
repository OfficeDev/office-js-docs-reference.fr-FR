| Classe | Champs | Description |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getcount--)|Obtient le nombre de liaisons de la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/excel/excel.bindingcollection#getitemornullobject-id-)|Obtient un objet de liaison par ID.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getcount--)|Renvoie le nombre de graphiques dans la feuille de calcul.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.chartcollection#getitemornullobject-name-)|Extrait un graphique à l’aide de son nom.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getcount--)|Renvoie le nombre de points de graphique dans la série.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getcount--)|Renvoie le nombre de séries de la collection.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|Spécifie le commentaire associé à ce nom.|
||[delete()](/javascript/api/excel/excel.nameditem#delete--)|Supprime le nom donné.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getrangeornullobject--)|Renvoie l’objet de plage qui est associé au nom.|
||[scope](/javascript/api/excel/excel.nameditem#scope)|Spécifie si le nom est d’étendue au workbook ou à une feuille de calcul spécifique.|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|Renvoie la feuille de calcul dans laquelle est inclus l’élément nommé.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetornullobject)|Renvoie la feuille de calcul dans laquelle l’élément nommé est d’étendue.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add(name: string, reference: Range \| string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|Ajoute un nouveau nom à la collection de l’étendue donnée.|
||[addFormulaLocal(name: string, formula: string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|Ajoute un nouveau nom à la collection de l’étendue donnée à l’aide des paramètres régionaux de l’utilisateur pour la formule.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|Obtient le nombre d’éléments nommés dans la collection.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|Obtient un `NamedItem` objet à l’aide de son nom.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|Obtient le nombre de tableaux croisés dynamiques de la collection.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|Obtient un tableau croisé dynamique par nom.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données.|
||[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|Renvoie la plage utilisée d’un objet de plage donné.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|Obtient le nombre `RangeView` d’objets de la collection.|
|[Paramètre](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|Supprime le paramètre.|
||[key](/javascript/api/excel/excel.setting#key)|Clé qui représente l’ID du paramètre.|
||[value](/javascript/api/excel/excel.setting#value)|Représente la valeur stockée pour ce paramètre.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add(key: string, value: string \| number \| boolean \| Date Array \| <any> \| any)](/javascript/api/excel/excel.settingcollection#add-key--value-)|Définit ou ajoute le paramètre spécifié dans le classeur.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|Obtient le nombre de paramètres de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|Obtient une entrée de paramètre via la clé.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|Obtient une entrée de paramètre via la clé.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|Se produit lorsque les paramètres du document sont modifiés.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[paramètres](/javascript/api/excel/excel.settingschangedeventargs#settings)|Obtient `Setting` l’objet qui représente la liaison qui a élevé l’événement de changement de paramètres|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|Obtient le nombre de tableaux de la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|Obtient un tableau à l’aide de son nom ou de son ID.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|Obtient le nombre de colonnes dans le tableau.|
||[getItemOrNullObject(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|Obtient un objet de colonne par son nom ou son ID.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|Obtient le nombre de lignes dans le tableau.|
|[Classeur](/javascript/api/excel/excel.workbook)|[paramètres](/javascript/api/excel/excel.workbook#settings)|Représente une collection de paramètres associés au workbook.|
|[Feuille de calcul](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|La plage utilisée est la plus petite plage qui englobe toutes les cellules auxquelles une valeur ou un format est affecté.|
||[noms](/javascript/api/excel/excel.worksheet#names)|Collection de noms inclus dans l’étendue de la feuille de calcul active.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount(visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|Obtient le nombre de feuilles de calcul dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID.|
