# <a name="outlook-add-in-api-requirement-set-12"></a>Ensemble de conditions requises de l’API du complément Outlook 1.2

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!NOTE]
> Cette documentation est une [exigence](/javascript/office/requirement-sets/outlook-api-requirement-sets) autre que le dernier jeu d’exigence. 

## <a name="whats-new-in-12"></a>Nouveautés de la version 1.2

L’ensemble de conditions requises de la version 1.2 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Désormais, les compléments peuvent insérer du texte au niveau du curseur de l’utilisateur, soit dans l’objet ou le corps du message.

### <a name="change-log"></a>Journal des modifications

- Ajouté [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): retourne de manière asynchrone les données sélectionnées à partir de l’objet ou le corps du message.
- Ajout de la méthode [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback) : Insère les données dans le corps ou l’objet d’un message de manière asynchrone.
- Modification de la fonction [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata) : Ajout de la propriété `attachments` dans le paramètre `formData`.
- Modification de la fonction [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata) : Ajout de la propriété `attachments` dans le paramètre `formData`.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](https://docs.microsoft.com/outlook/add-ins/quick-start)