| Classe | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Fournit des informations basées sur les paramètres de culture système actuels.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche de la virgule pour les valeurs numériques.|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Spécifie si les séparateurs système d’Excel sont activés.|
|[Commentaire](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Obtient les entités (par exemple, les personnes) mentionnées dans les commentaires.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Obtient le contenu des commentaires enrichis (par exemple, les mentions dans les commentaires).|
||[résolu](/javascript/api/excel/excel.comment#resolved)|État du thread de commentaire.|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Met à jour le contenu des commentaires avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Crée un nouveau commentaire avec le contenu donné sur la cellule donnée.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Adresse de messagerie de l’entité mentionnée dans un commentaire.|
||[id](/javascript/api/excel/excel.commentmention#id)|ID de l’entité.|
||[name](/javascript/api/excel/excel.commentmention#name)|Nom de l’entité mentionnée dans un commentaire.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Entités (par exemple, personnes) mentionnées dans les commentaires.|
||[résolu](/javascript/api/excel/excel.commentreply#resolved)|État de réponse du commentaire.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|Contenu de commentaire enrichi (par exemple, mentions dans les commentaires).|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Met à jour le contenu des commentaires avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Crée une réponse de commentaire pour un commentaire.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Tableau contenant toutes les entités (par exemple, les personnes) mentionnées dans le commentaire.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|Spécifie le contenu enrichi du commentaire (par exemple, le contenu de commentaire avec mentions, la première entité mentionnée a un attribut d’ID de 0 et la seconde entité mentionnée a un attribut d’ID de 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Obtient le nom de la culture au format languagecode2-country/regioncode2 (par exemple, « zh-cn » ou « en-us »).|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Définit le format adapté à la culture de l’affichage des nombres.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche de la virgule pour les valeurs numériques.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Range \| string)](/javascript/api/excel/excel.range#moveto-destinationrange-)|Déplace les valeurs, la mise en forme et les formules des cellules de la plage actuelle vers la plage de destination, en remplaçant les anciennes informations de ces cellules.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Ajuste le retrait de la mise en forme de plage.|
|[Classeur](/javascript/api/excel/excel.workbook)|[Fermer (closeBehavior ? : Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fermer le classeur actif.|
||[Enregistrer (saveBehavior ? : Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Enregistrer le classeur actif.|
|[Feuille de calcul](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Se produit lorsque l’état masqué d’une ou plusieurs lignes a changé dans une feuille de calcul spécifique.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Adresse de la plage qui a effectué le calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Se produit lorsque l’état masqué d’une ou plusieurs lignes a changé dans une feuille de calcul spécifique.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtient le type de modification qui représente la façon dont l’événement a été déclenché.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle les données ont été modifiées.|
