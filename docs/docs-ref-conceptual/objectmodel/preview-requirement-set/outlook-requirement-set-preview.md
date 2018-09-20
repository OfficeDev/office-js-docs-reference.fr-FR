# <a name="outlook-add-in-api-preview-requirement-set"></a>Ensemble de conditions requises de l’API du complément Outlook (aperçu)

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!NOTE]
> Cette documentation est pour un **Aperçu** [exigence](/javascript/office/requirement-sets/outlook-api-requirement-sets). Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions. Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément. La disponibilité des méthodes et des propriétés présentées dans cet ensemble de conditions doit être testée avant de les utiliser.

L’ensemble de l’aperçu exigence inclut toutes les fonctionnalités [d’exigence défini 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md).

## <a name="features-in-preview"></a>Fonctionnalités (aperçu) :

Les fonctionnalités suivantes sont disponibles en aperçu.

- [À partir de](/javascript/api/outlook/office.from) - ajouter un nouvel objet qui fournit une méthode pour obtenir la valeur affectée à.
- [Organisateur](/javascript/api/outlook/office.organizer) - Ajout d’un nouvel objet qui fournit une méthode pour obtenir la valeur de l’organisateur.
- [Périodicité](/javascript/api/outlook/office.recurrence) - Ajout d’un nouvel objet qui fournit des méthodes pour obtenir et définir la périodicité d’un rendez-vous, mais uniquement obtenir la périodicité des messages qui sont des demandes de réunion.
- [SeriesTime](/javascript/api/outlook/office.seriestime) - Ajout d’un nouvel objet qui fournit des méthodes pour obtenir et définir les dates et heures de rendez-vous dans une série périodique et pour obtenir les dates et heures de demandes de réunion dans une série périodique.
- [Event.Completed](/javascript/api/office/office.addincommands.event#completed-options-) : nouveau paramètre facultatif `options`, qui est un dictionnaire ayant comme seule valeur valide `allowEvent`. Cette valeur est utilisée pour annuler l’exécution d’un événement.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - Ajout d’une nouvelle méthode qui associe un fichier à partir de l’en base64 codage à un message ou un rendez-vous.
- [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback) - Ajout d’une nouvelle méthode qui ajoute un gestionnaire d’événements pour un événement pris en charge.
- [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) - modifié pour obtenir l’à partir de la valeur en mode composition.
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) - Ajout d’une fonction qui renvoie les données d’initialisation transmises quand le complément est [activé par un message actionnable](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) - en modifié pour obtenir la valeur de la bibliothèque multimédia en mode composition.
- [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) - ajouté une nouvelle propriété qui obtient ou définit un objet qui fournit des méthodes pour gérer la périodicité d’un élément de rendez-vous. Cette propriété peut également être utilisée pour obtenir la périodicité d’une réunion élément de demande.
- [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback) - Ajout d’une nouvelle méthode qui supprime un gestionnaire d’événements.
- [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string) - Ajout d’une nouvelle propriété qui obtient l’id de la série une occurrence appartient.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) - Ajout d’un accès à `getAccessTokenAsync`, qui permet aux compléments d’[obtenir un jeton d’accès](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) pour l’API Microsoft Graph.
- [Office.MailboxEnums.Days](/javascript/api/outlook/office.mailboxenums.days) - Ajout d’une nouvelle énumération qui spécifie le jour de semaine ou le type de la journée.
- [Office.MailboxEnums.Month](/javascript/api/outlook/office.mailboxenums.month) - Ajout d’une nouvelle énumération qui spécifie le mois.
- [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook/office.mailboxenums.recurrencetimezone) - Ajout d’une nouvelle énumération qui spécifie le fuseau horaire appliqué à la périodicité.
- [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook/office.mailboxenums.recurrencetype) - Ajout d’une nouvelle énumération qui spécifie le type de périodicité.
- [Office.MailboxEnums.WeekNumber](/javascript/api/outlook/office.mailboxenums.weeknumber) - Ajout d’une nouvelle énumération qui spécifie la semaine du mois.
- [Office.EventType](/javascript/api/office/office.eventtype) - modifié pour prendre en charge les événements RecurrenceChanged, RecipientsChanged, AppointmentTimeChanged et OfficeThemeChanged via l’ajout de `RecurrencePatternChanged`, `RecipientsChanged`, `AppointmentTimeChanged`, et `OfficeThemeChanged` entrées respectivement.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](https://docs.microsoft.com/outlook/add-ins/quick-start)