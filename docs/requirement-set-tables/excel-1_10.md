| Classe | Champs | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Contenu du commentaire.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Supprime le commentaire et toutes les réponses connectées.|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Obtient la cellule où se trouve ce commentaire.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Obtenir l’adresse email de l’auteur du commentaire.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Obtient le nom de l’auteur du commentaire.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Obtenir l’heure de création du commentaire.|
||[id](/javascript/api/excel/excel.comment#id)|Spécifie l’identificateur de commentaire.|
||[Réponses](/javascript/api/excel/excel.comment#replies)|Représente une collection de feuilles de calcul associées au classeur.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Crée un nouveau commentaire avec le contenu donné sur la cellule donnée.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Obtient le nombre de commentaires de la collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Obtient un commentaire à partir de la collection de sites en fonction de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Obtient un commentaire en fonction de sa position dans la collection.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Obtient le commentaire à partir de la cellule spécifiée.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Obtient le commentaire auquel la réponse donnée est connectée.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Contenu de la réponse au commentaire.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Supprime la réponse de commentaire.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Obtient la cellule où se trouve cette réponse de commentaire.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Obtient le commentaire parent de cette réponse.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Obtenir l’adresse email de l’auteur de la réponse au commentaire.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Obtenir le nom de l’auteur de la réponse au commentaire.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Obtient l’heure de création de la réponse à un commentaire.|
||[id](/javascript/api/excel/excel.commentreply#id)|Spécifie l’identificateur de réponse du commentaire.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Crée une réponse de commentaire pour un commentaire.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Obtient le nombre de réponses aux commentaires de la collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Renvoie une réponse de commentaire identifié via son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Obtient une réponse de commentaire en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Spécifie si la liste des champs peut être affichée dans l’interface utilisateur.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Supprime le style de tableau croisé dynamique.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Crée une copie de ce style de tableau croisé dynamique avec des copies de tous les éléments de style.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Obtient le nom du style de tableau croisé dynamique.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Spécifie si cet `PivotTableStyle` objet est en lecture seule.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Crée un vide `PivotTableStyle` avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Obtient le nombre de styles tableaux croisés dynamiques de la collection.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Obtient le style de tableau croisé dynamique par défaut pour l’étendue de l’objet parent.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Obtient `PivotTableStyle` une par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Obtient `PivotTableStyle` une par nom.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Définit le style de tableau croisé dynamique par défaut à utiliser dans l’étendue de l’objet parent.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#group-groupoption-)|Groupe les colonnes et les lignes d’un plan.|
||[hideGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|Masque les détails du groupe de lignes ou de colonnes.|
||[height](/javascript/api/excel/excel.range#height)|Renvoie la distance en points, pour un zoom à 100 %, entre le bord supérieur de la plage et le bord inférieur de la plage.|
||[left](/javascript/api/excel/excel.range#left)|Renvoie la distance en points, pour un zoom à 100 %, entre le bord gauche de la feuille de calcul et le bord gauche de la plage.|
||[top](/javascript/api/excel/excel.range#top)|Renvoie la distance en points, pour un zoom à 100 %, entre le bord supérieur de la feuille de calcul et le bord supérieur de la plage.|
||[width](/javascript/api/excel/excel.range#width)|Renvoie la distance en points, pour un zoom à 100 %, entre le bord gauche de la plage et le bord droit de la plage.|
||[showGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|Affiche les détails du groupe de lignes ou de colonnes.|
||[ungroup(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#ungroup-groupoption-)|Regroupe les colonnes et les lignes d’un plan.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Copie et copie un `Shape` objet.|
||[placement](/javascript/api/excel/excel.shape#placement)|Représente la manière dont l’objet est attaché aux cellules en dessous.|
|[Segment](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Représente la légende du slicer.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Supprime tous les filtres appliqués actuellement sur le tableau.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Supprime le segment.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Renvoie une matrice de noms d’éléments sélectionnés.|
||[height](/javascript/api/excel/excel.slicer#height)|Représente la hauteur, exprimée en points, de l’axe de graphique.|
||[left](/javascript/api/excel/excel.slicer#left)|Représente la distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[name](/javascript/api/excel/excel.slicer#name)|Représente le nom du slicer.|
||[id](/javascript/api/excel/excel.slicer#id)|Représente l’ID unique du slicer.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|La valeur est `true` si tous les filtres appliqués actuellement sur le slicer sont effacés.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Représente la collection d’éléments de slicer qui font partie du slicer.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Obtenir la feuille de calcul contenant la plage.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Sélectionne les éléments de slicer en fonction de leurs clés.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Représente l’ordre de tri des éléments dans le segment.|
||[style](/javascript/api/excel/excel.slicer#style)|Valeur constante qui représente le style de slicer.|
||[top](/javascript/api/excel/excel.slicer#top)|Représente la distance, en points, du bord supérieur de la section à la partie droite de la feuille de calcul.|
||[width](/javascript/api/excel/excel.slicer#width)|Représente la largeur, en points, de la forme.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[Ajouter (slicerSource : chaîne \| tableau croisé dynamique \| Table, sourceField : chaîne \| PivotField \| nombre \| TableColumn, slicerDestination ? : chaîne \| feuille de calcul)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Ajoute un nouveau segment au classeur.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Renvoie le nombre de séries de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Obtient un objet slicer à l’aide de son nom ou de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Obtient un slicer à l’aide de son nom ou de son ID.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|La valeur `true` est si l’élément de slicer est sélectionné.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|La valeur est `true` si l’élément de slicer possède des données.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Représente la valeur unique représentant l’élément de segment.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Représente le titre affiché dans l’interface utilisateur Excel.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Renvoie le nombre de segment les éléments dans le segment.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Obtient un segment objet de l’élément à l’aide de son nom ou clé.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Obtient un segment de l’élément à l’aide de son nom ou clé.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Supprime le style de slicer.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Crée une copie de ce style de slicer avec des copies de tous les éléments de style.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Obtient le nom du style de slicer.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Spécifie si cet `SlicerStyle` objet est en lecture seule.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Crée un style de slicer vide avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Obtient le nombre de styles de slicer de la collection.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Obtient la valeur `SlicerStyle` par défaut pour l’étendue de l’objet parent.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Obtient `SlicerStyle` une par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Obtient `SlicerStyle` une par nom.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Définit le style de slicer par défaut à utiliser dans l’étendue de l’objet parent.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Supprime le style de tableau.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Crée une copie de ce style de tableau avec des copies de tous les éléments de style.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Obtient le nom du style de tableau.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Spécifie si cet `TableStyle` objet est en lecture seule.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Crée un vide `TableStyle` avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Obtient le nombre de styles de tableaux de la collection.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Obtient le style de tableau par défaut pour l’étendue de l’objet parent.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Obtient `TableStyle` une par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Obtient `TableStyle` une par nom.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Définit le style de tableau par défaut à utiliser dans l’étendue de l’objet parent.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Supprime le style de tableau.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Crée une copie de ce style de chronologie avec des copies de tous les éléments de style.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Obtient le nom du style de chronologie.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Spécifie si cet `TimelineStyle` objet est en lecture seule.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Crée un vide `TimelineStyle` avec le nom spécifié.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Obtient le nombre de styles de délai de la collection.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Obtient le style de chronologie par défaut pour l’étendue de l’objet parent.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Obtient `TimelineStyle` une par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Obtient `TimelineStyle` une par nom.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[setDefault (newDefaultStyle : PivotTableStyle \| chaîne)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Définit le style de chronologie par défaut à utiliser dans l’étendue de l’objet parent.|
|[Classeur](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Obtient le segment actif actuel du classeur.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Obtient le segment actif actuel du classeur.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Représente une collection de commentaires associés au workbook.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Représente une collection de PivotTableStyles associée au classeur.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Représente une collection de styles associés au classeur.|
||[Slicers](/javascript/api/excel/excel.workbook#slicers)|Représente une collection de slicers associés au workbook.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Représente une collection de TableStyles associés au classeur.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Représente une collection de TimelineStyles associés au classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Renvoie une collection de tous les objets Lecteur sur l’ordinateur.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Se produit lorsqu’une ou plusieurs colonnes ont été triées.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Se produit lorsqu’une ou plusieurs lignes ont été triées.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Se produit lorsqu’une action clic gauche/clic se produit dans la feuille de calcul.|
||[Slicers](/javascript/api/excel/excel.worksheet#slicers)|Renvoie une collection de slicers qui font partie de la feuille de calcul.|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|Affiche les groupes de lignes ou de colonnes par niveaux de plan.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Se produit lorsqu’une ou plusieurs colonnes ont été triées.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Se produit lorsqu’une ou plusieurs lignes ont été triées.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Se produit lorsque l’opération clic gauche/clic se produit dans la collection de feuilles de calcul.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle le tri s’est produit.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Obtient l’adresse de plage qui représente les zones sélectionnées dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle le tri s’est produit.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[adresse](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Obtient l’adresse qui représente la cellule sur laquelle vous avez fait un clic gauche/appuyé pour une feuille de calcul spécifique.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|Distance, en points, entre le point de clic gauche et le point de clic gauche vers le bord gauche (ou de droite pour les langues qui s’viennent de droite à gauche) du quadrillage de la cellule cliquée/tapée à gauche.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|La distance en points à partir du point sur lequel vous avez effectué un clic gauche/avez appuyé vers le bord supérieur du bord de la grille de la cellule clic gauche/tape.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle la cellule a été cliquée/tapée avec le bouton gauche.|
