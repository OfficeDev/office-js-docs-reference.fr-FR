# <a name="outlook-add-in-api-requirement-set-14"></a>Ensemble de conditions requises de l’API du complément Outlook 1.4

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!NOTE]
> Cette documentation est une [exigence](/javascript/office/requirement-sets/outlook-api-requirement-sets) autre que le dernier jeu d’exigence.

## <a name="whats-new-in-14"></a>Nouveautés de la version 1.4

L’ensemble de conditions requises de la version 1.4 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Il comprend en plus l’accès à l’espace de noms `Office.ui`.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) : Affiche une boîte de dialogue dans un hôte Office.
- Ajout de la méthode [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-) : Remet un message de la part de la boîte de dialogue à sa page parent/d’ouverture.
- Objet [dialogue](/javascript/api/office/office.dialog) ajouté : l’objet qui est renvoyée lorsque le [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) méthode est appelée.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](https://docs.microsoft.com/outlook/add-ins/quick-start)