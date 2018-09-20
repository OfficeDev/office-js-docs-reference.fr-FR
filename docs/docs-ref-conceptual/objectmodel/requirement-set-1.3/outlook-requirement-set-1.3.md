# <a name="outlook-add-in-api-requirement-set-13"></a>Ensemble de conditions requises de l’API du complément Outlook 1.3

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!NOTE]
> Cette documentation est une [exigence](/javascript/office/requirement-sets/outlook-api-requirement-sets) autre que le dernier jeu d’exigence. 

## <a name="whats-new-in-13"></a>Nouveautés de la version 1.3

L’ensemble de conditions requises de la version 1.3 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). Les fonctionnalités suivantes ont été ajoutées :

- Prise en charge des [commandes de complément](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook).
- Possibilité d’enregistrer ou de fermer un élément en cours de composition.
- Amélioration de l’objet [Body](/javascript/api/outlook_1_3/office.body) pour autoriser les compléments à obtenir ou à définir la totalité du corps du message.
- Nouvelles méthodes de conversion pour convertir les ID aux formats EWS et REST.
- Possibilité d’ajouter des messages de notification dans la barre d’informations sur les éléments.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Body.getAsync](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) : Renvoie le corps actif dans un format spécifié.
- Ajout de la méthode [Body.setAsync](/javascript/api/outlook_1_3/office.body#setasync-data--options--callback-) : Remplace l’ensemble du corps avec le texte spécifié.
- Ajout de la propriété [Office.context.officeTheme](office.context.md#officetheme-object) : permet d’accéder aux couleurs du thème Office.
- Ajout de l’objet [Event](/javascript/api/office/office.addincommands.event) : transmis comme paramètre aux fonctions de commande sans IU dans un complément Outlook. Utilisé pour signaler la fin du traitement de l’événement.
- Ajout de la méthode [Office.context.mailbox.item.close](office.context.mailbox.item.md#close) : Ferme l’élément en cours qui est composé.
- Ajout de la méthode [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#saveasyncoptions-callback) : Enregistre un élément de manière asynchrone.
- Ajout de l’objet [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages) : Obtient les messages de notification pour un élément.
- Ajout de la méthode [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#converttoewsiditemid-restversion--string) : Convertit un ID d’élément mis en forme pour REST au format EWS.
- Ajout de la méthode [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) : Convertit un ID d’élément mis en forme pour EWS au format REST.
- Ajout de l’énumération [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook_1_3/office.mailboxenums.itemnotificationmessagetype) : Spécifie le type de message de notification pour un rendez-vous ou un message.
- Ajout de l’énumération [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion) : Spécifie la version de l’API REST qui correspond à un ID d’élément au format REST.
- Ajout de l’objet [NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages) : fournit des méthodes pour accéder aux messages de notification dans un complément Outlook.
- Ajout du type [NotificationMessageDetails](/javascript/api/outlook_1_3/office.notificationmessagedetails) : renvoyé par la méthode `NotificationMessages.getAllAsync`.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](https://docs.microsoft.com/outlook/add-ins/quick-start)