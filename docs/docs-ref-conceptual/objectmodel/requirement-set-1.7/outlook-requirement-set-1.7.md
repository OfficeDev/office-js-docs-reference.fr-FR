# <a name="outlook-add-in-api-requirement-set-17"></a>Outlook complément nécessite API définie 1.7

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

## <a name="whats-new-in-17"></a>Quelles sont les nouveautés dans 1.7 ?

Ensemble de conditions requises 1.7 inclut toutes les fonctionnalités [d’exigence défini 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md). Il ajouté les fonctionnalités suivantes.

- Ajout de nouvelles API concernant la périodicité des rendez-vous et des messages qui sont des demandes de réunion.
- Modifier la propriété item.from pour qu’il soit également disponible en mode composition.
- Prise en charge supplémentaire pour les événements RecurrenceChanged, RecipientsChanged et AppointmentTimeChanged.

### <a name="change-log"></a>Journal des modifications

- Ajoutée [à partir de](/javascript/api/outlook_1_7/office.from): ajoute un nouvel objet qui fournit une méthode pour obtenir la valeur affectée à.
- Ajout [d’organisateur](/javascript/api/outlook_1_7/office.organizer): ajoute un nouvel objet qui fournit une méthode pour obtenir la valeur de l’organisateur.
- Ajout de [périodicité](/javascript/api/outlook_1_7/office.recurrence): ajoute un nouvel objet qui fournit des méthodes pour obtenir et définir la périodicité d’un rendez-vous, mais obtenir uniquement la périodicité des messages sont les demandes de réunion.
- Ajouté [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone): ajoute un nouvel objet qui représente la configuration de fuseau horaire de la périodicité.
- Ajouté [SeriesTime](/javascript/api/outlook_1_7/office.seriestime): ajoute un nouvel objet qui fournit des méthodes pour obtenir et définir les dates et heures de rendez-vous dans une série périodique et pour obtenir les dates et heures de demandes de réunion dans une série périodique.
- Ajouté [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback): ajoute une nouvelle méthode qui ajoute un gestionnaire d’événements pour un événement pris en charge.
- Modifié [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom): modifie pour obtenir la valeur en mode composition.
- Modifié [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) - modifie pour obtenir la valeur de la bibliothèque multimédia en mode composition.
- Ajouté [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence): ajoute une nouvelle propriété qui obtient ou définit un objet qui fournit des méthodes pour gérer la périodicité d’un élément de rendez-vous. Cette propriété peut également être utilisée pour obtenir la périodicité d’une réunion élément de demande.
- Ajouté [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback): ajoute une nouvelle méthode qui supprime un gestionnaire d’événements.
- Ajouté [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string): ajoute une nouvelle propriété qui obtient l’id de la série une occurrence appartient.
- Ajouté [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days): ajoute une nouvelle énumération qui spécifie le jour de semaine ou le type de la journée.
- Ajouté [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month): ajoute une nouvelle énumération qui spécifie le mois.
- Ajouté [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone): ajoute une nouvelle énumération qui spécifie le fuseau horaire appliqué à la périodicité.
- Ajouté [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype): ajoute une nouvelle énumération qui spécifie le type de périodicité.
- Ajouté [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber): ajoute une nouvelle énumération qui spécifie la semaine du mois.
- Modifié [Office.EventType](/javascript/api/office/office.eventtype): modifie pour prendre en charge les événements RecurrenceChanged, RecipientsChanged et AppointmentTimeChanged via l’ajout de `RecurrenceChanged`, `RecipientsChanged`, et `AppointmentTimeChanged` entrées respectivement.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](https://docs.microsoft.com/outlook/add-ins/quick-start)