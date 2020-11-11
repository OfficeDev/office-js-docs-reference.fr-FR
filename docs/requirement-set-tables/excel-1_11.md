| Class | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Fournit des informations basées sur les paramètres de culture système actuels.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche du séparateur décimal pour les valeurs numériques.|
||[UseSystemSeparators,](/javascript/api/excel/excel.application#usesystemseparators)|Indique si les séparateurs système d’Excel sont activés.|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Obtient les entités (par exemple, les personnes) mentionnées dans les commentaires.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Obtient le contenu de commentaire enrichi (par exemple, mentions dans les commentaires).|
||[évaluation](/javascript/api/excel/excel.comment#resolved)|État du fil de commentaire.|
||[updateMentions (contentWithMentions : Excel. CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Met à jour le contenu de commentaire avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[Add (cellAddress : Range \| String, content : CommentRichContent \| String, ContentType ?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Crée un nouveau commentaire avec le contenu donné sur la cellule donnée.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Adresse de messagerie de l’entité mentionnée dans Comment.|
||[id](/javascript/api/excel/excel.commentmention#id)|ID de l’entité.|
||[name](/javascript/api/excel/excel.commentmention#name)|Nom de l’entité mentionnée dans Comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Entités (par exemple, les personnes) mentionnées dans les commentaires.|
||[évaluation](/javascript/api/excel/excel.commentreply#resolved)|État de la réponse de commentaire.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|Contenu de commentaire enrichi (par exemple, mentions dans les commentaires).|
||[updateMentions (contentWithMentions : Excel. CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Met à jour le contenu de commentaire avec une chaîne spécialement mise en forme et une liste de mentions.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[Add (Content : CommentRichContent \| String, ContentType ?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Crée une réponse à un commentaire pour un commentaire.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Tableau contenant toutes les entités (par exemple, les personnes) mentionnées dans le commentaire.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|Spécifie le contenu enrichi du commentaire (par exemple, le contenu du commentaire avec des mentions, la première entité mentionnée a un attribut ID de 0 et la deuxième entité mentionnée a un attribut ID égal à 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Obtient le nom de la culture au format languagecode2-Country/regioncode2 (par exemple, « zh-CN » ou « en-US »).|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Définit le format d’affichage des nombres approprié pour la culture.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche du séparateur décimal pour les valeurs numériques.|
|[Range](/javascript/api/excel/excel.range)|[moveTo (destinationRange : chaîne de plage \| )](/javascript/api/excel/excel.range#moveto-destinationrange-)|Déplace les valeurs de cellule, la mise en forme et les formules de la plage actuelle à la plage de destination, en remplaçant les anciennes informations de ces cellules.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent (montant : nombre)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Ajuste la mise en retrait de la plage de mise en forme.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Fermer (closeBehavior ? : Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fermer le classeur actif.|
||[Enregistrer (saveBehavior ? : Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Enregistrer le classeur actif.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Survient lorsque l’état masqué d’une ou plusieurs lignes a été modifié sur une feuille de calcul spécifique.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Adresse de la plage qui a terminé le calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Survient lorsque l’état masqué d’une ou plusieurs lignes a été modifié sur une feuille de calcul spécifique.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtient le type de modification qui représente la manière dont l’événement a été déclenché.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|
