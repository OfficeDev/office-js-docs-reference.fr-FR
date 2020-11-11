| Class | Champs | Description |
|:---|:---|:---|
|[Boîte aux lettres](/javascript/api/outlook/outlook.mailbox)|[addHandlerAsync (eventType : chaîne Office. EventType \| , Handler : (type : Office. EventType) => void, Options ?: Office. AsyncContextOptions, callback ?: (AsyncResult : Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#addhandlerasync-eventtype--handler--type-)|ajoute un gestionnaire d’événements pour un événement pris en charge.|
||[getCallbackTokenAsync (options : Office. AsyncContextOptions & {isRest ?: Boolean}, callback : (asyncResult : Office. AsyncResult <string> ) => void)](/javascript/api/outlook/outlook.mailbox#getcallbacktokenasync-options--isrest--callback--asyncresult-)|Obtient une chaîne qui contient un jeton servant à appeler des API REST ou des services Web Exchange (EWS).|
||[isRest](/javascript/api/outlook/outlook.mailbox#isrest)||
||[removeHandlerAsync (eventType : chaîne Office. EventType \| , Options ?: Office. AsyncContextOptions, callback ?: (AsyncResult : Office. AsyncResult <void> ) => void)](/javascript/api/outlook/outlook.mailbox#removehandlerasync-eventtype--options--callback--asyncresult-)|Supprime les gestionnaires d’événements pour un type d’événement pris en charge.|
||[restUrl](/javascript/api/outlook/outlook.mailbox#resturl)|obtient l’URL du point de terminaison REST de ce compte de messagerie.|
