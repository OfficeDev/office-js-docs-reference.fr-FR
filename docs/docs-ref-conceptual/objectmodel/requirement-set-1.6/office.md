 

# <a name="office"></a>Bureau

L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](/javascript/api/office).

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="members-and-methods"></a>Membres et méthodes

| Membre | Type |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Membre |
| [CoercionType](#coerciontype-string) | Membre |
| [EventType](#eventtype-string) | Membre |
| [SourceProperty](#sourceproperty-string) | Membre |

### <a name="namespaces"></a>Espaces de noms

[contexte](office.context.md): fournit des interfaces partagées d’espace de noms de contexte au bureau, l’API de compléments pour une utilisation dans l’API du complément Outlook.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.

### <a name="members"></a>Membres

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

Spécifie le résultat d’un appel asynchrone.

##### <a name="type"></a>Type :

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Succeeded`| String|L’appel a réussi.|
|`Failed`| String|L’appel n’a pas réussi.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

---

####  <a name="coerciontype-string"></a>CoercionType :String

Indique comment forcer le type des données retournées ou définies par la méthode appelée.

##### <a name="type"></a>Type :

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Html`| String|Demande que les données soient renvoyées au format HTML.|
|`Text`| String|Demande que les données soient renvoyées au format texte.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

---

####  <a name="eventtype-string"></a>EventType :String

spécifie l’événement associé à un gestionnaire d’événements.

##### <a name="type"></a>Type :

*   String

##### <a name="properties"></a>Propriétés :

| Nom | Type | Description |
|---|---|---|
|`ItemChanged`| String | L’élément sélectionné a changé. |

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1,5 |
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture |

---

####  <a name="sourceproperty-string"></a>SourceProperty :String

Spécifie la source des données renvoyées par la méthode appelée.

##### <a name="type"></a>Type :

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Body`| String|La source de données est dans le corps d’un message.|
|`Subject`| String|La source de données est dans l’objet d’un message.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|