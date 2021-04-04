| Classe | Champs | Description |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Spécifie l’angle vers lequel le texte est orienté pour le titre de l’axe du graphique.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtient les valeurs d’une dimension unique de la série de graphiques.|
|[Commentaire](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Obtient le type de contenu du commentaire.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Obtient `CommentDetail` le tableau qui contient l’ID de commentaire et les ID de ses réponses connexes.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Spécifie la source de l’événement.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle l’événement s’est produit.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Obtient le type de modification qui représente la façon dont l’événement modifié est déclenché.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Obtenez le `CommentDetail` tableau qui contient l’ID de commentaire et les ID de ses réponses connexes.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Spécifie la source de l’événement.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle l’événement s’est produit.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Se produit lorsque les commentaires sont ajoutés.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Se produit lorsque des commentaires ou des réponses dans une collection de commentaires sont modifiés, y compris lorsque les réponses sont supprimées.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Se produit lorsque des commentaires sont supprimés dans la collection de commentaires.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Obtient `CommentDetail` le tableau qui contient l’ID de commentaire et les ID de ses réponses connexes.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Spécifie la source de l’événement.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle l’événement s’est produit.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|Représente l’ID du commentaire.|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|Représente les ID des réponses associées qui appartiennent au commentaire.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|Type de contenu de la réponse.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Définit le format adapté à la culture de l’affichage de la date et de l’heure.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Obtient la chaîne utilisée comme séparateur de date.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Obtient la chaîne de format pour une valeur de date longue.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Obtient la chaîne de format pour une valeur de temps longue.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Obtient la chaîne de format pour une valeur de date courte.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Obtient la chaîne utilisée comme séparateur d’heure.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparator](/javascript/api/excel/excel.pivotdatefilter#comparator)|Le comparateur est la valeur statique à laquelle les autres valeurs sont comparées.|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|Spécifie la condition du filtre, qui définit les critères de filtrage nécessaires.|
||[exclusive](/javascript/api/excel/excel.pivotdatefilter#exclusive)|Si `true` , le filtre exclut *les* éléments qui répondent aux critères.|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|Limite inférieure de la plage pour la `between` condition de filtre.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|Limite supérieure de la plage pour la `between` condition de filtre.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|Pour `equals` , et les conditions de `before` `after` `between` filtre, indique si les comparaisons doivent être réalisées en tant que jours entiers.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel.PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Définit un ou plusieurs des filtres de tableau croisé dynamique actuels du champ et les applique au champ.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Permet d’effacer tous les critères de tous les filtres du champ.|
||[clearFilter(filterType: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Permet d’effacer tous les critères existants du filtre du champ du type donné (s’il en existe un actuellement appliqué).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|Obtient tous les filtres actuellement appliqués sur le champ.|
||[isFiltered(filterType?: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Vérifie s’il existe des filtres appliqués sur le champ.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|Filtre de date actuellement appliqué au champ de tableau croisé dynamique.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|Filtre d’étiquette actuellement appliqué au champ de tableau croisé dynamique.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|Filtre manuel actuellement appliqué au champ de tableau croisé dynamique.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|Filtre de valeurs actuellement appliqué au champ de tableau croisé dynamique.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparator](/javascript/api/excel/excel.pivotlabelfilter#comparator)|Le comparateur est la valeur statique à laquelle les autres valeurs sont comparées.|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|Spécifie la condition du filtre, qui définit les critères de filtrage nécessaires.|
||[exclusive](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|Si `true` , le filtre exclut *les* éléments qui répondent aux critères.|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|Limite inférieure de la plage pour la `between` condition de filtre.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|Sous-stration utilisée pour `beginsWith` , et les conditions de `endsWith` `contains` filtre.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|Limite supérieure de la plage pour la `between` condition de filtre.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Liste des éléments sélectionnés à filtrer manuellement.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Spécifie si le tableau croisé dynamique autorise l’application de plusieurs filtres de tableau croisé dynamique sur un champ de tableau croisé dynamique donné dans le tableau.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Obtient le nombre de tableaux croisés dynamiques dans la collection.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Obtient le premier tableau croisé dynamique de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Obtient un tableau croisé dynamique par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Obtient un tableau croisé dynamique par nom.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparator](/javascript/api/excel/excel.pivotvaluefilter#comparator)|Le comparateur est la valeur statique à laquelle les autres valeurs sont comparées.|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|Spécifie la condition du filtre, qui définit les critères de filtrage nécessaires.|
||[exclusive](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Si `true` , le filtre exclut *les* éléments qui répondent aux critères.|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|Limite inférieure de la plage pour la `between` condition de filtre.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Spécifie si le filtre est pour les éléments N supérieur/inférieur, le pourcentage N supérieur/inférieur ou la somme N supérieure/inférieure.|
||[seuil](/javascript/api/excel/excel.pivotvaluefilter#threshold)|Nombre seuil « N » d’éléments, de pourcentage ou de somme à filtrer pour une condition de filtre supérieure/inférieure.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|Limite supérieure de la plage pour la `between` condition de filtre.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Nom de la « valeur » choisie dans le champ par lequel filtrer.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|Renvoie un objet qui représente la plage contenant tous les antécédents directs d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
||[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Obtient une collection étendue de tableaux croisés dynamiques qui chevauchent la plage.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtient l’objet de plage contenant la cellule d’ancrage de la cellule dans laquelle la cellule est étendue.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Représente si toutes les cellules ont une bordure renversée.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Représente la catégorie du format de nombre de chaque cellule.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Représente si toutes les cellules sont enregistrées en tant que formule ma matrice.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|Obtient le nombre `RangeAreas` d’objets de cette collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|Renvoie `RangeAreas` l’objet en fonction de la position dans la collection.|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Renvoie `RangeAreas` l’objet en fonction de l’ID de feuille de calcul ou du nom de la collection.|
||[getRangeAreasOrNullObjectBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Renvoie `RangeAreas` l’objet en fonction du nom de la feuille de calcul ou de l’ID de la collection.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Renvoie un tableau d’adresses de style A1.|
||[Zones](/javascript/api/excel/excel.workbookrangeareas#areas)|Renvoie `RangeAreasCollection` l’objet.|
||[plages](/javascript/api/excel/excel.workbookrangeareas#ranges)|Renvoie les plages qui composent cet objet dans un  `RangeCollection`   objet.|
|[Feuille de calcul](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Obtient une collection de propriétés personnalisées au niveau de la feuille de calcul.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Supprime la propriété personnalisée.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Obtient la clé de la propriété personnalisée.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Obtient ou définit la valeur de la propriété personnalisée.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Ajoute une nouvelle propriété personnalisée qui s’ajoute à la clé fournie.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Obtient le nombre de propriétés personnalisées dans cette feuille de calcul.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
