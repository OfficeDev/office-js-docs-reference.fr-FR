| Classe | Champs | Description |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutid)|Spécifie l’ID d’une mise en page des diapositives à utiliser pour la nouvelle diapositive.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slidemasterid)|Spécifie l’ID d’un curseur de diapositive à utiliser pour la nouvelle diapositive.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slidemasters)|Renvoie la collection `SlideMaster` d’objets qui se retrouvent dans la présentation.|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#tags)|Renvoie une collection de balises attachées à la présentation.|
|[Forme](/javascript/api/powerpoint/powerpoint.shape)|[id](/javascript/api/powerpoint/powerpoint.shape#id)|Obtient l’ID unique de la forme.|
||[tags](/javascript/api/powerpoint/powerpoint.shape#tags)|Renvoie une collection de balises dans la forme.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getcount--)|Obtient le nombre de formes dans la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitem-key-)|Obtient une forme à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemat-index-)|Obtient une forme à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemornullobject-id-)|Obtient une forme à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[disposition](/javascript/api/powerpoint/powerpoint.slide#layout)|Obtient la mise en page de la diapositive.|
||[Formes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Renvoie une collection de formes dans la diapositive.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slidemaster)|Obtient `SlideMaster` l’objet qui représente le contenu par défaut de la diapositive.|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|Renvoie une collection de balises dans la diapositive.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add-options-)|Ajoute une nouvelle diapositive à la fin de la collection.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Obtient l’ID unique de la mise en page des diapositives.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Obtient le nom de la mise en page des diapositives.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getcount--)|Obtient le nombre de dispositions dans la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitem-key-)|Obtient une disposition à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemat-index-)|Obtient une disposition à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemornullobject-id-)|Obtient une disposition à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Obtient l’ID unique du curseur de diapositive.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Obtient la collection de mises en page fournies par le maître des diapositives pour les diapositives.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Obtient le nom unique du curseur de diapositive.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getcount--)|Obtient le nombre de diapositives de la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitem-key-)|Obtient un curseur de diapositive à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemat-index-)|Obtient un curseur de diapositive à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemornullobject-id-)|Obtient un curseur de diapositive à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|Obtient l’ID unique de la balise.|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|Obtient la valeur de la balise.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add-key--value-)|Ajoute une nouvelle balise à la fin de la collection.|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete-key-)|Supprime la balise avec la balise donnée `key` dans cette collection.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getcount--)|Obtient le nombre de balises dans la collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitem-key-)|Obtient une balise à l’aide de son ID unique.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemat-index-)|Obtient une balise à l’aide de son index de base zéro dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemornullobject-key-)|Obtient une balise à l’aide de son ID unique.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
