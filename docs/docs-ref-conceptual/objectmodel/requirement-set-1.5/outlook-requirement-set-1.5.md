# <a name="outlook-add-in-api-requirement-set-15"></a>Ensemble de conditions requises de l’API du complément Outlook 1.5

Le sous-ensemble Outlook-dans l’API de l’API JavaScript pour Office inclut des objets, méthodes, propriétés et événements que vous pouvez utiliser dans Outlook complément.

> [!NOTE]
> Cette documentation est une [exigence](/javascript/office/requirement-sets/outlook-api-requirement-sets) autre que le dernier jeu d’exigence.

## <a name="whats-new-in-15"></a>Nouveautés de la version 1.5

L’ensemble de conditions requises de la version 1.5 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). Les fonctionnalités suivantes ont été ajoutées :

- Prise en charge des [volets Office épinglables](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).
- Prise en charge de l’appel des [API REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api).
- Possibilité de marquer une pièce jointe comme élément incorporé.
- Possibilité de fermer un volet Office ou une boîte de dialogue.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback) : ajoute un gestionnaire d’événements pour un événement pris en charge.
- Ajouté [Office.EventType](office.md#eventtype-string): Spécifie l’événement associé à un gestionnaire d’événements et prend en charge l’événement ItemChanged.
- Ajout de la propriété [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string) : obtient l’URL du point de terminaison REST de ce compte de messagerie.
- Modification de la méthode [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback) : cette nouvelle version comprend une nouvelle signature (`getCallbackTokenAsync([options], callback)`). La version d’origine est toujours disponible et reste inchangée.
- Ajout de [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).
- Modification de la méthode [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback) : nouvelle valeur du dictionnaire `options` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.
- Modification de la fonction [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata) : nouvelle valeur du dictionnaire `formData.attachments` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.
- Modification de la fonction [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata) : nouvelle valeur du dictionnaire `formData.attachments` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](https://docs.microsoft.com/outlook/add-ins/quick-start)