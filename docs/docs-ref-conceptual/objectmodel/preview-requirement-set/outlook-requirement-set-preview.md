# <a name="outlook-add-in-api-preview-requirement-set"></a>Ensemble de conditions requises de l’API du complément Outlook (aperçu)

Le sous-ensemble Outlook-dans l’API de l’API JavaScript pour Office inclut des objets, méthodes, propriétés et événements que vous pouvez utiliser dans Outlook complément.

> [!NOTE]
> Cette documentation est pour un **Aperçu** [exigence](/javascript/office/requirement-sets/outlook-api-requirement-sets). Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions. Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément. La disponibilité des méthodes et des propriétés présentées dans cet ensemble de conditions doit être testée avant de les utiliser.

L’ensemble de l’aperçu exigence inclut toutes les fonctionnalités [d’exigence défini 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Fonctionnalités (aperçu) :

Les fonctionnalités suivantes sont disponibles en aperçu.

- [SharedProperties](/javascript/api/outlook/office.sharedproperties) - Ajout d’un nouvel objet qui représente les propriétés d’un élément de rendez-vous ou de message dans un dossier partagé, le calendrier ou la boîte aux lettres.
- [Event.Completed](/javascript/api/office/office.addincommands.event#completed-options-) : nouveau paramètre facultatif `options`, qui est un dictionnaire ayant comme seule valeur valide `allowEvent`. Cette valeur est utilisée pour annuler l’exécution d’un événement.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - Ajout d’une nouvelle méthode qui associe un fichier à partir de l’en base64 codage à un message ou un rendez-vous.
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) - Ajout d’une fonction qui renvoie les données d’initialisation transmises quand le complément est [activé par un message actionnable](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback) - Ajout d’une nouvelle méthode qui obtient un objet qui représente le sharedProperties d’un rendez-vous ou un élément de message.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) - Ajout d’un accès à `getAccessTokenAsync`, qui permet aux compléments d’[obtenir un jeton d’accès](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) pour l’API Microsoft Graph.
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) - Ajout d’un nouvel indicateur enum bits qui spécifie les autorisations de délégué.
- [Office.EventType](/javascript/api/office/office.eventtype) - modifié pour prendre en charge les événements OfficeThemeChanged grâce à l’ajout de `OfficeThemeChanged` entrée.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](https://docs.microsoft.com/outlook/add-ins/quick-start)