| Class | Champs | Description |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[addHandlerAsync (eventType : chaîne Office. EventType \| , Handler : any, Options ?: Office. AsyncContextOptions, callback ?: (AsyncResult : Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|ajoute un gestionnaire d’événements pour un événement pris en charge.|
||[organizer](/javascript/api/outlook/outlook.appointmentcompose#organizer)|Obtient l’organisateur de la réunion spécifiée.|
||[instances](/javascript/api/outlook/outlook.appointmentcompose#recurrence)|Obtient ou définit la périodicité d’un rendez-vous.|
||[removeHandlerAsync (eventType : chaîne Office. EventType \| , Options ?: Office. AsyncContextOptions, callback ?: (AsyncResult : Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentcompose#removehandlerasync-eventtype--options--callback--asyncresult-)|Supprime les gestionnaires d’événements pour un type d’événement pris en charge.|
||[seriesId](/javascript/api/outlook/outlook.appointmentcompose#seriesid)|Obtient l’ID de la série à laquelle une instance appartient.|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[addHandlerAsync (eventType : chaîne Office. EventType \| , Handler : any, Options ?: Office. AsyncContextOptions, callback ?: (AsyncResult : Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|ajoute un gestionnaire d’événements pour un événement pris en charge.|
||[instances](/javascript/api/outlook/outlook.appointmentread#recurrence)|Obtient la périodicité d’un rendez-vous.|
||[removeHandlerAsync (eventType : chaîne Office. EventType \| , Options ?: Office. AsyncContextOptions, callback ?: (AsyncResult : Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.appointmentread#removehandlerasync-eventtype--options--callback--asyncresult-)|Supprime les gestionnaires d’événements pour un type d’événement pris en charge.|
||[seriesId](/javascript/api/outlook/outlook.appointmentread#seriesid)|Obtient l’ID de la série à laquelle une instance appartient.|
|[AppointmentTimeChangedEventArgs](/javascript/api/outlook/outlook.appointmenttimechangedeventargs)|[end](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#end)||
||[start](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#start)||
||[type](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#type)||
|[From](/javascript/api/outlook/outlook.from)|[getAsync (options ?: Office. AsyncContextOptions, callback ?: (asyncResult : Office. AsyncResult <EmailAddressDetails> ) => void)](/javascript/api/outlook/outlook.from#getasync-options--callback--asyncresult-)|Obtient la valeur de l’d’un message.|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[addHandlerAsync (eventType : chaîne Office. EventType \| , Handler : any, Options ?: Office. AsyncContextOptions, callback ?: (AsyncResult : Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messagecompose#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|ajoute un gestionnaire d’événements pour un événement pris en charge.|
||[from](/javascript/api/outlook/outlook.messagecompose#from)|Obtient l’adresse de messagerie de l’expéditeur d’un message.|
||[removeHandlerAsync (eventType : chaîne Office. EventType \| , Options ?: Office. AsyncContextOptions, callback ?: (AsyncResult : Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messagecompose#removehandlerasync-eventtype--options--callback--asyncresult-)|Supprime les gestionnaires d’événements pour un type d’événement pris en charge.|
||[seriesId](/javascript/api/outlook/outlook.messagecompose#seriesid)|Obtient l’ID de la série à laquelle une instance appartient.|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[addHandlerAsync (eventType : chaîne Office. EventType \| , Handler : any, Options ?: Office. AsyncContextOptions, callback ?: (AsyncResult : Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messageread#addhandlerasync-eventtype--handler--options--callback--asyncresult-)|ajoute un gestionnaire d’événements pour un événement pris en charge.|
||[instances](/javascript/api/outlook/outlook.messageread#recurrence)|Obtient la périodicité d’un rendez-vous.|
||[removeHandlerAsync (eventType : chaîne Office. EventType \| , Options ?: Office. AsyncContextOptions, callback ?: (AsyncResult : Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.messageread#removehandlerasync-eventtype--options--callback--asyncresult-)|Supprime les gestionnaires d’événements pour un type d’événement pris en charge.|
||[seriesId](/javascript/api/outlook/outlook.messageread#seriesid)|Obtient l’ID de la série à laquelle une instance appartient.|
|[Organizer](/javascript/api/outlook/outlook.organizer)|[getAsync (options ?: Office. AsyncContextOptions, callback ?: (asyncResult : Office. AsyncResult <EmailAddressDetails> ) => void)](/javascript/api/outlook/outlook.organizer#getasync-options--callback--asyncresult-)|Obtient la valeur de l’organisateur d’un rendez-vous sous la forme {@link Office. EmailAddressDetails | Objet EmailAddressDetails}|
|[RecipientsChangedEventArgs](/javascript/api/outlook/outlook.recipientschangedeventargs)|[changedRecipientFields](/javascript/api/outlook/outlook.recipientschangedeventargs#changedrecipientfields)||
||[type](/javascript/api/outlook/outlook.recipientschangedeventargs#type)||
|[RecipientsChangedFields](/javascript/api/outlook/outlook.recipientschangedfields)|[bcc](/javascript/api/outlook/outlook.recipientschangedfields#bcc)|Obtient une valeur qui indique si le champ **CCI** a été modifié.|
||[cc](/javascript/api/outlook/outlook.recipientschangedfields#cc)|Obtient si les destinataires dans le champ **CC** ont été modifiés.|
||[optionalAttendees](/javascript/api/outlook/outlook.recipientschangedfields#optionalattendees)|Obtient si les participants facultatifs ont été modifiés.|
||[requiredAttendees](/javascript/api/outlook/outlook.recipientschangedfields#requiredattendees)|Obtient si les participants requis ont été modifiés.|
||[resources](/javascript/api/outlook/outlook.recipientschangedfields#resources)|Obtient si les ressources ont été modifiées.|
||[to](/javascript/api/outlook/outlook.recipientschangedfields#to)|Obtient si les destinataires dans le champ **à** ont été modifiés.|
|[Périodicité](/javascript/api/outlook/outlook.recurrence)|[getAsync (options ?: Office. AsyncContextOptions, callback ?: (asyncResult : Office. AsyncResult <Recurrence> ) => void)](/javascript/api/outlook/outlook.recurrence#getasync-options--callback--asyncresult-)|Renvoie l’objet de périodicité actuelle d’une série de rendez-vous.|
||[recurrenceProperties](/javascript/api/outlook/outlook.recurrence#recurrenceproperties)|Obtient ou définit les propriétés de la série de rendez-vous périodiques.|
||[recurrenceTimeZone](/javascript/api/outlook/outlook.recurrence#recurrencetimezone)|Obtient ou définit les propriétés de la série de rendez-vous périodiques.|
||[recurrenceType](/javascript/api/outlook/outlook.recurrence#recurrencetype)|Obtient ou définit le type de série de rendez-vous périodiques.|
||[seriesTime](/javascript/api/outlook/outlook.recurrence#seriestime)|Le {@link Office. SeriesTime | L’objet SeriesTime} vous permet de gérer les dates de début et de fin de la série de rendez-vous périodiques et|
||[setAsync (recurrencePattern : recurrence, Options ?: Office. AsyncContextOptions, callback ?: (asyncResult : Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.recurrence#setasync-recurrencepattern--options--callback--asyncresult-)|Définit la périodicité d’une série de rendez-vous.|
|[RecurrenceChangedEventArgs](/javascript/api/outlook/outlook.recurrencechangedeventargs)|[instances](/javascript/api/outlook/outlook.recurrencechangedeventargs#recurrence)||
||[type](/javascript/api/outlook/outlook.recurrencechangedeventargs#type)||
|[RecurrenceProperties](/javascript/api/outlook/outlook.recurrenceproperties)|[dayOfMonth](/javascript/api/outlook/outlook.recurrenceproperties#dayofmonth)|Représente le jour du mois.|
||[dayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#dayofweek)|Représente le jour de la semaine ou le type de jour, par exemple, jour de week-end vs semaine.|
||[chômés](/javascript/api/outlook/outlook.recurrenceproperties#days)|Représente l’ensemble des jours de cette périodicité.|
||[firstDayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#firstdayofweek)|Représente le premier jour de la semaine choisi, autrement la valeur par défaut est celle des paramètres de l’utilisateur actuel.|
||[interval](/javascript/api/outlook/outlook.recurrenceproperties#interval)|Représente la période entre les instances de la même série périodique.|
||[Mois](/javascript/api/outlook/outlook.recurrenceproperties#month)|Représente le mois.|
||[weekNumber](/javascript/api/outlook/outlook.recurrenceproperties#weeknumber)|Représente le numéro de la semaine dans le mois sélectionné, par exemple, « premier » pour la première semaine du mois.|
|[RecurrenceTimeZone](/javascript/api/outlook/outlook.recurrencetimezone)|[name](/javascript/api/outlook/outlook.recurrencetimezone#name)|Représente le nom du fuseau horaire de périodicité.|
||[compensé](/javascript/api/outlook/outlook.recurrencetimezone#offset)|Valeur entière représentant la différence, en minutes, entre le fuseau horaire local et l’heure UTC à la date de début de la série de réunions.|
|[SeriesTime](/javascript/api/outlook/outlook.seriestime)|[getDuration()](/javascript/api/outlook/outlook.seriestime#getduration--)|Obtient la durée en minutes d’une instance habituelle dans une série de rendez-vous périodiques.|
||[getEndDate()](/javascript/api/outlook/outlook.seriestime#getenddate--)|Obtient la date de fin d’une périodicité dans les éléments suivants|
||[getEndTime()](/javascript/api/outlook/outlook.seriestime#getendtime--)|Obtient l’heure de fin d’une instance de rendez-vous ou de demande de réunion usuelle d’une périodicité dans le fuseau horaire de l’utilisateur ou|
||[getStartDate()](/javascript/api/outlook/outlook.seriestime#getstartdate--)|Obtient la date de début d’une périodicité dans l’exemple suivant|
||[getStartTime()](/javascript/api/outlook/outlook.seriestime#getstarttime--)|Obtient l’heure de début d’une instance de rendez-vous classique d’une périodicité dans le fuseau horaire dans lequel l’utilisateur/complément a défini le paramètre|
||[setDuration (minutes : nombre)](/javascript/api/outlook/outlook.seriestime#setduration-minutes-)|Définit la durée de tous les rendez-vous dans une périodicité.|
||[setEndDate (date : chaîne)](/javascript/api/outlook/outlook.seriestime#setenddate-date-)|Définit la date de fin d’une série de rendez-vous périodiques.|
||[setEndDate (Year : Number, month : Number, Day : Number)](/javascript/api/outlook/outlook.seriestime#setenddate-year--month--day-)|Définit la date de fin d’une série de rendez-vous périodiques.|
||[setStartDate (date : chaîne)](/javascript/api/outlook/outlook.seriestime#setstartdate-date-)|Définit la date de début d’une série de rendez-vous périodiques.|
||[setStartDate (Year : Number, month : Number, Day : Number)](/javascript/api/outlook/outlook.seriestime#setstartdate-year--month--day-)|Définit la date de début d’une série de rendez-vous périodiques.|
||[setStartTime (heures : nombre, minutes : nombre)](/javascript/api/outlook/outlook.seriestime#setstarttime-hours--minutes-)|Définit l’heure de début de toutes les instances d’une série de rendez-vous périodiques selon le fuseau horaire défini pour le critère de périodicité.|
||[setStartTime (Time : chaîne)](/javascript/api/outlook/outlook.seriestime#setstarttime-time-)|Définit l’heure de début de toutes les instances d’une série de rendez-vous périodiques selon le fuseau horaire défini pour le critère de périodicité.|
