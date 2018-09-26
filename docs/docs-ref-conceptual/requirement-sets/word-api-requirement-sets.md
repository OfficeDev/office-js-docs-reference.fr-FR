# <a name="word-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Word

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Compléments Office utilisent ensembles spécifiés dans le manifeste ou une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API nécessitant un complément. Pour plus d’informations, voir [définit les versions d’Office et de la spécification](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Compléments Word s’exécuter sur plusieurs versions d’Office, notamment Office 2016 ou version ultérieure pour Windows, Office pour iPad, Office pour Mac et Office Online. Le tableau suivant répertorie les ensembles Word, les applications hôtes Office qui prennent en charge que les numéros set et la version ou la version requise pour les applications.

> [!NOTE]
> Pour ensembles sont marqués comme étant bêta, utilisez la version spécifiée (ou version ultérieure) du logiciel Office et la bibliothèque de la version bêta du CDN : https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> 
> Entrées ne pas répertoriées comme bêta sont généralement disponibles et vous pouvez continuer à utiliser la bibliothèque de Production CDN :https://appsforoffice.microsoft.com/lib/1/hosted/office.js

|  Ensemble de conditions requises  |   Office 365 pour Windows\*  |  Office 365 pour iPad  |  Office 365 pour Mac  | Office Online  | Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| WordApi 1.3 | Version 1612 (Build 7668.1000) ou version ultérieure| Mars 2017, 2.22 ou version ultérieure | Mars 2017, 15.32 ou version ultérieure| Mars 2017 ||
| WordApi 1.2  | Mise à jour de décembre 2015, version 1601 (Build 6568.1000) ou version ultérieure | Janvier 2016, 1.18 ou version ultérieure | Janvier 2016, 15.19 ou version ultérieure| Septembre 2016 | |
| WordApi 1.1  | Version 1509 (Build 4266.1001) ou version ultérieure| Janvier 2016, 1.18 ou version ultérieure | Janvier 2016, 15.19 ou version ultérieure| Septembre 2016 | |

> [!NOTE]
> Le numéro de version pour 2016 Office installée via MSI est 16.0.4266.1001. Cette version contient uniquement l’ensemble de conditions requises WordApi 1.1.

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

- [Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver le numéro de version et de build pour une application cliente Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Présentation d’Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="whats-new-in-word-javascript-api-13"></a>Nouveautés de l’API JavaScript 1.3 pour Word 

Les ajouts apportés aux API JavaScript pour Word dans l’ensemble de conditions requises 1.3 sont présentés ci-dessous. 

|Objet| Nouveautés| Description|Ensemble de conditions requises| 
|:-----|-----|:----|:----| 
|[application](/javascript/api/word/word.application)|_Méthode_ > createDocument(base64File: string) | Crée un nouveau document à partir d’un fichier .docx encodé en base 64. En lecture seule.|1.3|
|[body](/javascript/api/word/word.body)|_Relation_ > lists|Obtient la collection d’objets list dans le corps. En lecture seule.|1.3|
|[body](/javascript/api/word/word.body)|_Relation_ > parentBody|Obtient le corps parent du corps. Par exemple, le corps parent du corps d’une cellule de tableau peut être un en-tête. En lecture seule.|1.3|
|[body](/javascript/api/word/word.body)|_Relation_ > parentSection|Obtient la section parent du corps. En lecture seule.|1.3|
|[body](/javascript/api/word/word.body)|_Relation_ > styleBuiltIn|Obtient ou définit le nom de style intégré pour le corps. Utilisez cette propriété pour les styles intégrés qui sont portables entre les paramètres régionaux. Pour utiliser des styles personnalisés ou des noms de style localisés, consultez la propriété « style ».|1.3|
|[body](/javascript/api/word/word.body)|_Relation_ > tables|Obtient la collection d’objets table dans le corps. En lecture seule.|1.3|
|[body](/javascript/api/word/word.body)|_Relation_ > type|Obtient le type du corps. Le type peut être « MainDoc », « Section », « Header », « Footer » ou « TableCell ». En lecture seule.|1.3|
|[body](/javascript/api/word/word.body)|_Méthode_ > getRange(rangeLocation: RangeLocation)|Obtient la totalité du corps, ou le point de début ou de fin du corps, sous la forme d’une plage.|1.3|
|[body](/javascript/api/word/word.body)|_Méthode_ > insertTable (rowCount : nombre, columnCount : nombre, insertLocation : InsertLocation, valeurs : chaîne)|Insère un tableau avec le nombre spécifié de lignes et de colonnes. La valeur insertLocation peut être « Start » (début) ou « End » (fin).|1.3|
|[breaktype](/javascript/api/word/word.breaktype)|_Relation_ > sauts|Spécifie la forme d’un saut : ligne, page ou type de section. En lecture seule.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relation_ > lists|Obtient la collection d’objets list du contrôle de contenu. En lecture seule.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relation_ > parentBody|Obtient le corps parent du contrôle de contenu. En lecture seule.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relation_ > parentTable|Obtient le tableau qui contient le contrôle de contenu. Renvoie un objet null s’il n’est pas contenu dans un tableau. En lecture seule.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relation_ > parentTableCell|Obtient la cellule de tableau qui contient le contrôle de contenu. Renvoie un objet null s’il n’est pas contenu dans une cellule de tableau. En lecture seule.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relation_ > styleBuiltIn|Obtient ou définit le nom de style intégré pour le contrôle de contenu. Utilisez cette propriété pour les styles intégrés qui sont portables entre les paramètres régionaux. Pour utiliser des styles personnalisés ou des noms de style localisés, consultez la propriété « style ».|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relation_ > subtype|Obtient le sous-type du contrôle de contenu. Le sous-type peut être « RichTextInline », « RichTextParagraphs », « RichTextTableCell », « RichTextTableRow » et « RichTextTable » pour les contrôles de contenu en texte enrichi. En lecture seule.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Relation_ > tables|Obtient la collection d’objets table du contrôle de contenu. En lecture seule.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Méthode_ > getRange(rangeLocation: RangeLocation)|Obtient le contrôle de contenu entier, ou le point  de début ou de fin du contrôle de contenu, sous la forme d’une plage.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Méthode_ > getTextRanges (endingMarks : string, trimSpacing : bool)|Obtient les plages du texte du contrôle de contenu en utilisant des signes de ponctuation et/ou d’autres marques de fin.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Méthode_ > insertTable (rowCount : nombre, columnCount : nombre, insertLocation : InsertLocation, valeurs : chaîne)|Insère un tableau avec le nombre spécifié de lignes et de colonnes dans un contrôle de contenu ou à proximité de celui-ci. La valeur insertLocation peut être définie sur « Start » (début), « End » (fin), « Before » (avant) ou « After » (après).|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Méthode_ > Fractionner (délimiteurs : chaîne, multiParagraphs : bool, trimDelimiters : bool, trimSpacing : bool)|Fractionne le contrôle de contenu en plages enfants à l’aide de délimiteurs.|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_Méthode_ > getByTypes(types: ContentControlType)|Obtient les contrôles de contenu qui ont les types et sous-types spécifiés.|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_Méthode_ > getFirst()|Obtient le premier contrôle de contenu de cette collection.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Propriété_ > key|Obtient la clé de la propriété personnalisée. En lecture seule. |1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Propriété_ > value|Obtient ou définit la valeur de la propriété personnalisée.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Relation_ > type|Obtient le type de valeur de la propriété personnalisée. En lecture seule.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Méthode_ > delete()|Supprime la propriété personnalisée.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Propriété_ > items|Collection d’objets customProperty. En lecture seule.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Méthode_ > deleteAll()|Supprime toutes les propriétés personnalisées de cette collection.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Méthode_ > getCount()|Obtient le nombre des propriétés personnalisées.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Méthode_ > getItem(key: string)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Méthode_ > définir (clé : string, valeur : objet)|Crée ou définit une propriété personnalisée.|1.3|
|[document](/javascript/api/word/word.document)|_Relation_ > properties|Obtient les propriétés du document actif. En lecture seule.|1.3|
|[document](/javascript/api/word/word.document)|_Méthode_ > open()|Ouvre le document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > applicationName|Obtient le nom de l’application du document. En lecture seule.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > author|Obtient ou définit l’auteur du document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > category|Obtient ou définit la catégorie du document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > comments|Obtient ou définit les commentaires du document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > company|Obtient ou définit la société du document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > format|Obtient ou définit le format du document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > keywords|Obtient ou définit les mots clés du document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > lastAuthor|Obtient ou définit le dernier auteur du document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > manager|Obtient ou définit le responsable du document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > revisionNumber|Obtient le numéro de révision du document. En lecture seule.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > security|Obtient la sécurité du document. En lecture seule.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > subject|Obtient ou définit le sujet du document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > template|Obtient le modèle du document. En lecture seule.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Propriété_ > title|Obtient ou définit le titre du document.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relation_ > creationDate|Obtient la date de création du document. En lecture seule.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relation_ > customProperties|Obtient la collection de propriétés personnalisées du document en lecture seule.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relation_ > lastPrintDate|Obtient la date dernière impression du document. En lecture seule.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Relation_ > lastSaveTime|Obtient le dernier enregistrement du document. En lecture seule.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Relation_ > parentTable|Obtient le tableau qui contient l’image insérée. Renvoie un objet null s’il n’est pas contenu dans un tableau. En lecture seule.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Relation_ > parentTableCell|Obtient la cellule de tableau qui contient l’image insérée. Renvoie un objet null s’il n’est pas contenu dans une cellule de tableau. En lecture seule.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Méthode_ > getNext()|Obtient l’image insérée suivante.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Méthode_ > getRange(rangeLocation: RangeLocation)|Obtient l’image, ou le point de départ ou de fin de l’image, sous la forme d’une plage.|1.3|
|[inlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|_Méthode_ > getFirst()|Obtient la première image insérée de cette collection.|1.3|
|[list](/javascript/api/word/word.list)|_Propriété_ > id|Obtient l’ID de la liste. En lecture seule.|1.3|
|[list](/javascript/api/word/word.list)|_Propriété_ > levelExistences|Vérifie si chacun des 9 niveaux existe dans la liste. Une valeur True indique le niveau existe, ce qui signifie qu’il existe au moins un élément de liste à ce niveau. En lecture seule.|1.3|
|[list](/javascript/api/word/word.list)|_Relation_ > levelTypes|Obtient les 9 types de niveau de la liste. Chaque type peut être « Bullet », « Number » ou « Picture ». En lecture seule.|1.3|
|[list](/javascript/api/word/word.list)|_Relation_ > paragraphs|Obtient les paragraphes de la liste. En lecture seule.|1.3|
|[list](/javascript/api/word/word.list)|_Méthode_ > getLevelParagraphs(level: number)|Obtient les paragraphes qui figurent au niveau spécifié de la liste.|1.3|
|[list](/javascript/api/word/word.list)|_Méthode_ > getLevelString(level: number)|Obtient la puce, le nombre ou l’image au niveau spécifié, sous la forme d’une chaîne.|1.3|
|[list](/javascript/api/word/word.list)|_Méthode_ > insertParagraph (façon : string, insertLocation : InsertLocation)|Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Start » (début), « End » (fin), « Before » (avant) ou « After » (après).|1.3|
|[list](/javascript/api/word/word.list)|_Méthode_ > setLevelAlignment (niveau : nombre, alignement : Alignment)|Définit l’alignement de la puce, du numéro ou de l’image au niveau spécifié de la liste.|1.3|
|[list](/javascript/api/word/word.list)|_Méthode_ > setLevelBullet (niveau : nombre, listBullet : ListBullet, le code de caractère : fontName number, : chaîne)|Définit le format de puce au niveau spécifié de la liste. Si la puce est « Custom », la valeur charCode est requise.|1.3|
|[list](/javascript/api/word/word.list)|_Méthode_ > setLevelIndents (niveau : nombre, textIndent : float, textIndent : float)|Définit les deux retraits du niveau spécifié de la liste.|1.3|
|[list](/javascript/api/word/word.list)|_Méthode_ > setLevelNumbering (niveau : nombre, listNumbering : ListNumbering, formatString : objet)|Définit le format de numérotation du niveau spécifié de la liste.|1.3|
|[list](/javascript/api/word/word.list)|_Méthode_ > setLevelStartingNumber (niveau : nombre, startingNumber : nombre)|Définit le numéro de départ du niveau spécifié de la liste. La valeur par défaut est 1.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Propriété_ > items|Collection d’objets de liste. En lecture seule.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Méthode_ > getById(id: number)|Obtient une liste par son identificateur.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Méthode_ > getFirst()|Obtient la première liste de cette collection.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Méthode_ > getItem(index: number)|Obtient un objet de liste en fonction de son indice dans la collection.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Propriété_ > level|Obtient ou définit le niveau de l’élément dans la liste.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Propriété_ > listString|Obtient la puce, le numéro ou l’image de l’élément de liste sous forme de chaîne. En lecture seule.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Propriété_ > siblingIndex|Obtient le numéro d’ordre de l’élément de liste relativement à ses frères. En lecture seule.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Méthode_ > getAncestor(parentOnly: bool)|Obtient le parent de l’élément de liste ou son ancêtre le plus proche si le parent n’existe pas.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Méthode_ > getDescendants(directChildrenOnly: bool)|Obtient tous les éléments de liste descendants de l’élément de liste.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Propriété_ > isLastParagraph|Indique que le paragraphe est le dernier au sein de son corps parent. En lecture seule.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Propriété_ > isListItem|Vérifie si le paragraphe est un élément de liste. En lecture seule.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Propriété_ > tableNestingLevel|Obtient le niveau de tableau du paragraphe. Renvoie 0 si le paragraphe ne figure pas dans un tableau. En lecture seule.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relation_ > list|Obtient la liste à laquelle ce paragraphe appartient. Renvoie un objet null si le paragraphe n’est pas dans une liste. En lecture seule.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relation_ > listItem|Obtient l’élément de liste du paragraphe. Renvoie un objet null si le paragraphe ne fait pas partie d’une liste. En lecture seule.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relation_ > parentBody|Obtient le corps parent du paragraphe. En lecture seule.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relation_ > parentTable|Obtient le tableau qui contient le paragraphe. Renvoie un objet null s’il n’est pas contenu dans un tableau. En lecture seule.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relation_ > parentTableCell|Obtient la cellule de tableau qui contient le paragraphe. Renvoie un objet null s’il n’est pas contenu dans une cellule de tableau. En lecture seule.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Relation_ > styleBuiltIn|Obtient ou définit le nom du style prédéfini du paragraphe. Utilisez cette propriété pour les styles intégrés qui sont portables entre les paramètres régionaux. Pour utiliser des styles personnalisés ou des noms de style localisés, consultez la propriété « style ».|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Méthode_ > attachToList (listId : nombre, niveau : nombre)|Permet au paragraphe de rejoindre une liste existante au niveau spécifié. Échoue si le paragraphe ne peut pas rejoindre la liste ou s’il est déjà un élément de liste.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Méthode_ > detachFromList()|Déplace ce paragraphe en dehors de la liste, si le paragraphe est un élément de liste.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Méthode_ > getNext()|Obtient le paragraphe suivant.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Méthode_ > getPrevious()|Obtient le paragraphe précédent.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Méthode_ > getRange(rangeLocation: RangeLocation)|Obtient le paragraphe entier, ou le point de début ou de fin du paragraphe, sous la forme d’une plage.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Méthode_ > getTextRanges (endingMarks : string, trimSpacing : bool)|Obtient les plages de texte du paragraphe au moyen de signes de ponctuation et/ou d’autres marques de fin.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Méthode_ > insertTable (rowCount : nombre, columnCount : nombre, insertLocation : InsertLocation, valeurs : chaîne)|Insère un tableau avec le nombre spécifié de lignes et de colonnes. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Méthode_ > Fractionner (délimiteurs : chaîne, trimDelimiters : bool, trimSpacing : bool)|Divise le paragraphe en plages enfants à l’aide de délimiteurs.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Méthode_ > startNewList()|Démarre une nouvelle liste avec ce paragraphe. Échoue si le paragraphe est déjà un élément de liste.|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_Méthode_ > getFirst()|Obtient le premier paragraphe de cette collection.|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_Méthode_ > getLast()|Obtient le dernier paragraphe dans cette collection.|1.3|
|[range](/javascript/api/word/word.range)|_Propriété_ > hyperlink|Obtient le premier lien hypertexte de la plage ou définit un lien hypertexte sur la plage. Tous les liens hypertexte de la plage sont supprimés lorsque vous définissez un nouveau lien hypertexte sur celle-ci. Utilisez un caractère de saut de ligne (« \n ») pour séparer la partie de l’adresse de la partie d’emplacement facultative.|1.3|
|[range](/javascript/api/word/word.range)|_Propriété_ > isEmpty|Vérifie si la longueur de la plage est zéro. En lecture seule.|1.3|
|[range](/javascript/api/word/word.range)|_Relation_ > lists|Obtient la collection d’objets de liste figurant dans la plage. En lecture seule.|1.3|
|[range](/javascript/api/word/word.range)|_Relation_ > parentBody|Obtient le corps parent de la plage. En lecture seule.|1.3|
|[range](/javascript/api/word/word.range)|_Relation_ > parentTable|Obtient le tableau qui contient la plage. Renvoie la valeur Null si elle n’est pas contenue dans un tableau. En lecture seule.|1.3|
|[range](/javascript/api/word/word.range)|_Relation_ > parentTableCell|Obtient la cellule de tableau qui contient la plage. Renvoie un objet null s’il n’est pas contenu dans une cellule de tableau. En lecture seule.|1.3|
|[range](/javascript/api/word/word.range)|_Relation_ > styleBuiltIn|Obtient ou définit le nom du style prédéfini de la plage. Utilisez cette propriété pour les styles intégrés qui sont portables entre les paramètres régionaux. Pour utiliser des styles personnalisés ou des noms de style localisés, consultez la propriété « style ».|1.3|
|[range](/javascript/api/word/word.range)|_Relation_ > tables|Obtient la collection d’objets de table dans la plage. En lecture seule.|1.3|
|[range](/javascript/api/word/word.range)|_Méthode_ > compareLocationWith(range: Range)|Compare l’emplacement de la plage à celui d’une autre plage.|1.3|
|[range](/javascript/api/word/word.range)|_Méthode_ > expandTo(range: Range)|Renvoie une nouvelle plage qui s’étend dans les deux directions à partir de cette plage pour couvrir une autre plage. Cette plage n’est pas modifiée.|1.3|
|[range](/javascript/api/word/word.range)|_Méthode_ > getHyperlinkRanges()|Obtient les plages enfants d’un lien hypertexte au sein de la plage.|1.3|
|[range](/javascript/api/word/word.range)|_Méthode_ > getNextTextRange (endingMarks : string, trimSpacing : bool)|Obtient la plage de texte suivante à l’aide de signes de ponctuation et/ou d’autres marques de fin.|1.3|
|[range](/javascript/api/word/word.range)|_Méthode_ > getRange(rangeLocation: RangeLocation)|Clone la plage ou obtient le point de début ou de fin de la plage sous la forme d’une nouvelle plage.|1.3|
|[range](/javascript/api/word/word.range)|_Méthode_ > getTextRanges (endingMarks : string, trimSpacing : bool)|Obtient les plages enfants de texte dans la plage à l’aide de signes de ponctuation et/ou d’autres marques de fin.|1.3|
|[range](/javascript/api/word/word.range)|_Méthode_ > insertTable (rowCount : nombre, columnCount : nombre, insertLocation : InsertLocation, valeurs : chaîne)|Insère un tableau avec le nombre spécifié de lignes et de colonnes. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|1.3|
|[range](/javascript/api/word/word.range)|_Méthode_ > intersectWith(range: Range)|Retourne une nouvelle plage en tant qu’intersection de cette plage avec une autre. Cette plage n’est pas modifiée.|1.3|
|[range](/javascript/api/word/word.range)|_Méthode_ > Fractionner (délimiteurs : chaîne, multiParagraphs : bool, trimDelimiters : bool, trimSpacing : bool)|Divise la plage en plages enfant à l’aide de délimiteurs.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Propriété_ > items|Collection d’objets de plage. En lecture seule.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Méthode_ > getFirst()|Obtient la première plage de cette collection.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Méthode_ > getItem(index: number)|Obtient un objet de plage en fonction de son indice dans la collection.|1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_Méthode_ > charger (objet : objet, option : objet)|Insère l’objet de proxy créé dans le calque JavaScript avec les propriétés et les options spécifiées dans le paramètre. |1.3|
|[requestContext](/javascript/api/word/word.requestcontext)|_Méthode_ > sync()|Envoie les demandes en file d’attente à Word et renvoie un objet de promesse, qui peut être utilisé pour ajouter d’autres actions en chaîne.|1.3|
|[section](/javascript/api/word/word.section)|_Méthode_ > getNext()|Obtient la section suivante.|1.3|
|[sectionCollection](/javascript/api/word/word.sectioncollection)|_Méthode_ > getFirst()|Obtient la première section de cette collection.|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > headerRowCount|Obtient et définit le nombre de lignes d’en-tête.|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > height|Obtient la hauteur du tableau en points. En lecture seule.|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > isUniform|Indique si toutes les lignes du tableau sont uniformes. En lecture seule.|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > nestingLevel|Obtient le niveau d’imbrication du tableau. Les tableaux de niveau supérieur ont le niveau 1. En lecture seule.|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > rowCount|Obtient le nombre de lignes dans le tableau. En lecture seule.|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > shadingColor|Obtient et définit la couleur d’ombrage.|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > style|Obtient ou définit le nom de style du tableau. Utilisez cette propriété pour les noms des styles personnalisés et localisés. Pour utiliser les styles prédéfinis qui sont portables entre différents paramètres régionaux, voir la propriété « styleBuiltIn ».|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > styleBandedColumns|Obtient et définit l’information qui indique que le tableau comporte des colonnes à bandes.|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > styleBandedRows|Obtient et définit l’information qui indique que le tableau comporte des lignes à bandes.|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > styleFirstColumn|Obtient et définit l’information qui indique si le tableau comporte une première colonne avec un style spécial.|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > styleLastColumn|Obtient et définit l’information qui indique si le tableau comporte une dernière colonne avec un style spécial.|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > styleTotalRow|Obtient et définit l’information qui indique si le tableau comporte une ligne de total (dernière ligne) avec un style spécial.|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > values|Obtient et définit les valeurs de texte du tableau, sous la forme d’un tableau Javascript 2D.|1.3|
|[table](/javascript/api/word/word.table)|_Propriété_ > width|Obtient et définit la largeur du tableau en points.|1.3|
|[table](/javascript/api/word/word.table)|_Relation_ > font|Obtient la police. Utilisez cette propriété pour obtenir et définir le nom de la police, la taille, la couleur et d’autres propriétés. En lecture seule.|1.3|
|[table](/javascript/api/word/word.table)|_Relation_ > horizontalAlignment|Obtient et définit l’alignement horizontal de chaque cellule du tableau. La valeur peut être « left » (gauche), « centered » (centré), « right » (droite) ou « justified » (justifié).|1.3|
|[table](/javascript/api/word/word.table)|_Relation_ > paragraphAfter|Obtient le paragraphe après le tableau. En lecture seule.|1.3|
|[table](/javascript/api/word/word.table)|_Relation_ > paragraphBefore|Obtient le paragraphe avant le tableau. En lecture seule.|1.3|
|[table](/javascript/api/word/word.table)|_Relation_ > parentBody|Obtient le corps parent du tableau. En lecture seule.|1.3|
|[table](/javascript/api/word/word.table)|_Relation_ > parentContentControl|Obtient le contrôle de contenu qui contient le tableau. En lecture seule.|1.3|
|[table](/javascript/api/word/word.table)|_Relation_ > parentTable|Obtient le tableau qui contient ce tableau. Renvoie un objet null s’il n’est pas contenu dans un tableau. En lecture seule.|1.3|
|[table](/javascript/api/word/word.table)|_Relation_ > parentTableCell|Obtient la cellule de tableau qui contient ce tableau. Renvoie un objet null s’il n’est pas contenu dans une cellule de tableau. En lecture seule.|1.3|
|[table](/javascript/api/word/word.table)|_Relation_ > rows|Obtient toutes les lignes du tableau. En lecture seule.|1.3|
|[table](/javascript/api/word/word.table)|_Relation_ > styleBuiltIn|Obtient ou définit le nom du style prédéfini du tableau. Utilisez cette propriété pour les styles intégrés qui sont portables entre les paramètres régionaux. Pour utiliser des styles personnalisés ou des noms de style localisés, consultez la propriété « style ».|1.3|
|[table](/javascript/api/word/word.table)|_Relation_ > tables|Obtient les tableaux enfants imbriqués au niveau de profondeur suivant. En lecture seule.|1.3|
|[table](/javascript/api/word/word.table)|_Relation_ > verticalAlignment|Obtient et définit l’alignement vertical de chaque cellule du tableau. La valeur peut être « top », « center » ou « bottom ».|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > addColumns (insertLocation : InsertLocation, columnCount : valeurs numériques : chaîne)|Ajoute des colonnes au début ou à la fin du tableau, en utilisant la première ou la dernière colonne existante en tant que modèle. Applicable aux tableaux uniformes. Si spécifiées, les valeurs de chaîne sont définies sur les lignes nouvellement insérées.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > addRows (insertLocation : InsertLocation, rowCount : valeurs numériques : chaîne)|Ajoute des lignes au début ou à la fin du tableau, en utilisant la première ou la dernière ligne existante en tant que modèle. Si spécifiées, les valeurs de chaîne sont définies sur les lignes nouvellement insérées.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > autoFitContents()|Ajuste automatiquement les colonnes du tableau à la largeur de leur contenu.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > autoFitWindow()|Ajuste automatiquement les colonnes du tableau à la largeur de la fenêtre.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > clear()|Efface le contenu du tableau.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > delete()|Supprime le tableau entier.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > deleteColumns (columnIndex : nombre, columnCount : nombre)|Supprime des colonnes spécifiques. Applicable aux tableaux uniformes.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > deleteRows (rowIndex : nombre, rowCount : nombre)|Supprime des lignes spécifiques.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > distributeColumns()|Répartit uniformément les largeurs de colonne.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > distributeRows()|Répartit uniformément les hauteurs de ligne.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > getBorder(borderLocation: BorderLocation)|Obtient le style de bordure de la bordure spécifiée.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > getCell (rowIndex : nombre, IndexCellule : nombre)|Obtient la cellule du tableau correspondant à une ligne et une colonne spécifiées.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Obtient la marge intérieure des cellules en points.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > getNext()|Obtient le tableau suivant.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > getRange(rangeLocation: RangeLocation)|Obtient la plage qui contient ce tableau, ou la plage située au début ou à la fin du tableau.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > insertContentControl()|Insère un contrôle de contenu dans le tableau.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > insertParagraph (façon : string, insertLocation : InsertLocation)|Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > insertTable (rowCount : nombre, columnCount : nombre, insertLocation : InsertLocation, valeurs : chaîne)|Insère un tableau avec le nombre spécifié de lignes et de colonnes. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > recherche (searchText : chaîne, searchOptions : ParamTypeStrings.SearchOptions)|Effectue une recherche avec les options de recherche spécifiées sur l’étendue de l’objet de tableau. Les résultats de la recherche sont un ensemble d’objets de plage.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > select(selectionMode: SelectionMode)|Sélectionne le tableau ou la position de début ou de fin du tableau et y accède dans l’interface utilisateur de Word.|1.3|
|[table](/javascript/api/word/word.table)|_Méthode_ > setCellPadding (cellPaddingLocation : CellPaddingLocation, cellPadding : float)|Définit la marge intérieure des cellules en points.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Propriété_ > color|Obtient ou définit la couleur de bordure de tableau, sous forme de valeur hexadécimale ou de nom.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Propriété_ > width|Obtient ou définit la largeur, en points, de la bordure du tableau. Non applicable aux types de bordure de tableau qui ont une largeur fixe.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Relation_ > type|Obtient ou définit le type de bordure du tableau.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propriété_ > cellIndex|Obtient l’index de la cellule dans la ligne correspondante. En lecture seule.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propriété_ > columnWidth|Obtient et définit la largeur de colonne de la cellule en points. Applicable aux tableaux uniformes.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propriété_ > rowIndex|Obtient l’index de la ligne de la cellule dans le tableau. En lecture seule.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propriété_ > shadingColor|Obtient ou définit la couleur d’ombrage de la cellule. La couleur est spécifiée au format « #RRVVBB » ou par son nom de couleur.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propriété_ > value|Obtient et définit le texte de la cellule.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Propriété_ > width|Obtient la largeur de la cellule en points. En lecture seule.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relation_ > body|Renvoie l’objet corps de la cellule. En lecture seule.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relation_ > horizontalAlignment|Obtient et définit l’alignement horizontal de la cellule. La valeur peut être « left » (gauche), « centered » (centré), « right » (droite) ou « justified » (justifié).|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relation_ > parentRow|Obtient la ligne parent de la cellule. En lecture seule.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relation_ > parentTable|Obtient le tableau parent de la cellule. En lecture seule.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Relation_ > verticalAlignment|Obtient et définit l’alignement vertical de la cellule. La valeur peut être « top », « center » ou « bottom ».|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Méthode_ > deleteColumn()|Supprime la colonne qui contient cette cellule. Applicable aux tableaux uniformes.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Méthode_ > deleteRow()|Supprime la ligne qui contient cette cellule.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Méthode_ > getBorder(borderLocation: BorderLocation)|Obtient le style de bordure de la bordure spécifiée.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Méthode_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Obtient la marge intérieure des cellules en points.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Méthode_ > getNext()|Obtient la cellule suivante.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Méthode_ > insertColumns (insertLocation : InsertLocation, columnCount : valeurs numériques : chaîne)|Ajoute des colonnes à gauche ou à droite de la cellule, en utilisant la colonne de la cellule en tant que modèle. Applicable aux tableaux uniformes. Si spécifiées, les valeurs de chaîne sont définies sur les lignes nouvellement insérées.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Méthode_ > insertRows (insertLocation : InsertLocation, rowCount : valeurs numériques : chaîne)|Insère les lignes au-dessus ou au-dessous de la cellule, en utilisant la ligne de la cellule en tant que modèle. Si spécifiées, les valeurs de chaîne sont définies sur les lignes nouvellement insérées.|1.3|
|[tableCell](/javascript/api/word/word.tablecell)|_Méthode_ > setCellPadding (cellPaddingLocation : CellPaddingLocation, cellPadding : float)|Définit la marge intérieure des cellules en points.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Propriété_ > items|Collection d’objets tableCell. En lecture seule.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Méthode_ > getFirst()|Obtient la première cellule de tableau de cette collection.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Méthode_ > getItem(index: number)|Obtient un objet de cellule de tableau en fonction de son index dans la collection.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Propriété_ > items|Collection d’objets de tableau. En lecture seule.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Méthode_ > getFirst()|Obtient le premier tableau de cette collection.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Méthode_ > getItem(index: number)|Obtient un objet de tableau en fonction de son index dans la collection.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propriété_ > cellCount|Obtient le nombre de cellules dans la ligne. En lecture seule.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propriété_ > isHeader|Vérifie si la ligne est une ligne d’en-tête. En lecture seule. Pour définir le nombre de lignes d’en-tête, utilisez HeaderRowCount sur l’objet de table. En lecture seule.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propriété_ > preferredHeight|Obtient et définit la hauteur de ligne préférée en points.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propriété_ > rowIndex|Obtient l’index de la ligne dans le tableau parent correspondant. En lecture seule.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propriété_ > shadingColor|Obtient et définit la couleur d’ombrage.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Propriété_ > values|Obtient et définit les valeurs de texte de la ligne, sous la forme d’un tableau Javascript 1D.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relation_ > cells|Obtient les cellules. En lecture seule.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relation_ > font|Obtient la police. Utilisez cette propriété pour obtenir et définir le nom de la police, la taille, la couleur et d’autres propriétés. En lecture seule.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relation_ > horizontalAlignment|Obtient et définit l’alignement horizontal de chaque cellule de la ligne. La valeur peut être « left » (gauche), « centered » (centré), « right » (droite) ou « justified » (justifié).|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relation_ > parentTable|Obtient la table parente. En lecture seule.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Relation_ > verticalAlignment|Obtient et définit l’alignement vertical des cellules de la ligne. La valeur peut être « top », « center » ou « bottom ».|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Méthode_ > clear()|Efface le contenu de la ligne.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Méthode_ > delete()|Supprime la ligne entière.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Méthode_ > getBorder(borderLocation: BorderLocation)|Obtient le style de bordure des cellules de la ligne.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Méthode_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Obtient la marge intérieure des cellules en points.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Méthode_ > getNext()|Obtient la ligne suivante.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Méthode_ > insertRows (insertLocation : InsertLocation, rowCount : valeurs numériques : chaîne)|Insère des lignes en utilisant cette ligne en tant que modèle. Si les valeurs sont spécifiées, insère les valeurs sur de nouvelles lignes.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Méthode_ > recherche (searchText : chaîne, searchOptions : ParamTypeStrings.SearchOptions)|Effectue une recherche avec les options de recherche spécifiées dans l’étendue de la ligne. Les résultats de la recherche sont un ensemble d’objets de plage.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Méthode_ > select(selectionMode: SelectionMode)|Sélectionne la ligne et y accède via l’interface utilisateur de Word.|1.3|
|[tableRow](/javascript/api/word/word.tablerow)|_Méthode_ > setCellPadding (cellPaddingLocation : CellPaddingLocation, cellPadding : float)|Définit la marge intérieure des cellules en points.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Propriété_ > items|Collection d’objets tableRow. En lecture seule.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Méthode_ > getFirst()|Obtient la première ligne de cette collection.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Méthode_ > getItem(index: number)|Obtient un objet de ligne de tableau en fonction de son index dans la collection.|1.3|


## <a name="whats-new-in-word-javascript-api-12"></a>Nouveautés de l’API JavaScript 1.2 pour Word

Les ajouts apportés aux API JavaScript pour Word dans l’ensemble de conditions requises 1.2 sont présentés ci-dessous. 

|Objet| Nouveautés| Description|Ensemble de conditions requises|
|:-----|-----|:----|:----|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Méthode_ > insertInlinePictureFromBase64 (base64EncodedImage : string, insertLocation : InsertLocation)|Insère une image incluse dans le contrôle de contenu, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Relation_ > paragraph|Obtient le paragraphe parent qui contient l’image insérée. En lecture seule.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Méthode_ > delete()|Supprime l’image insérée du document.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Méthode_ > insertBreak (breakType : BreakType, insertLocation : InsertLocation)|Insère un saut à l’emplacement spécifié du document principal. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Méthode_ > insertFileFromBase64 (base64File : string, insertLocation : InsertLocation)|Insère un document à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Méthode_ > insertHtml (html : string, insertLocation : InsertLocation)|Insère du code HTML à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Méthode_ > insertInlinePictureFromBase64 (base64EncodedImage : string, insertLocation : InsertLocation|Insère une image insérée à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Replace » (remplacer), « Before » (avant) ou « After » (après).|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Méthode_ > insertOoxml (ooxml : string, insertLocation : InsertLocation)|Insère du code OOXML à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Méthode_ > insertParagraph (façon : string, insertLocation : InsertLocation)|Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Méthode_ > insertText (texte : string, insertLocation : InsertLocation)|Insère du texte à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Méthode_ > select(selectionMode: SelectionMode)|Sélectionne l’image insérée. Word fait défiler le document jusqu’à accéder à la sélection.|1.2|
|[range](/javascript/api/word/word.range)|_Relation_ > inlinePictures|Obtient la collection d’objets image insérée de la plage. En lecture seule.|1.2|
|[range](/javascript/api/word/word.range)|_Méthode_ > insertInlinePictureFromBase64 (base64EncodedImage : string, insertLocation : InsertLocation)|Insère une image à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Replace » (remplacer), « Start » (début), « End » (fin), « Before » (avant) ou « After » (après).|1.2|

## <a name="word-javascript-api-11"></a>API JavaScript 1.1 pour Word

L’API JavaScript 1.1 pour Word est la première version de l’API. Pour plus d’informations à propos de l’API, voir la rubrique de référence sur l’[API JavaScript pour Word](/javascript/api/word). 

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des exigences en matière d’hôtes Office et d’API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
