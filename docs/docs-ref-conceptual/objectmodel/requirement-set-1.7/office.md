 

# <a name="office"></a><span data-ttu-id="92ed6-101">Bureau</span><span class="sxs-lookup"><span data-stu-id="92ed6-101">Office</span></span>

<span data-ttu-id="92ed6-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="92ed6-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="92ed6-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92ed6-104">Requirements</span></span>

|<span data-ttu-id="92ed6-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92ed6-105">Requirement</span></span>| <span data-ttu-id="92ed6-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="92ed6-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="92ed6-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92ed6-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92ed6-108">1.0</span><span class="sxs-lookup"><span data-stu-id="92ed6-108">1.0</span></span>|
|[<span data-ttu-id="92ed6-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92ed6-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="92ed6-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92ed6-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="92ed6-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="92ed6-111">Members and methods</span></span>

| <span data-ttu-id="92ed6-112">Membre</span><span class="sxs-lookup"><span data-stu-id="92ed6-112">Member</span></span> | <span data-ttu-id="92ed6-113">Type</span><span class="sxs-lookup"><span data-stu-id="92ed6-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="92ed6-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="92ed6-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="92ed6-115">Membre</span><span class="sxs-lookup"><span data-stu-id="92ed6-115">Member</span></span> |
| [<span data-ttu-id="92ed6-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="92ed6-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="92ed6-117">Membre</span><span class="sxs-lookup"><span data-stu-id="92ed6-117">Member</span></span> |
| [<span data-ttu-id="92ed6-118">EventType</span><span class="sxs-lookup"><span data-stu-id="92ed6-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="92ed6-119">Membre</span><span class="sxs-lookup"><span data-stu-id="92ed6-119">Member</span></span> |
| [<span data-ttu-id="92ed6-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="92ed6-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="92ed6-121">Membre</span><span class="sxs-lookup"><span data-stu-id="92ed6-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="92ed6-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="92ed6-122">Namespaces</span></span>

<span data-ttu-id="92ed6-123">[contexte](office.context.md): fournit des interfaces partagées d’espace de noms de contexte au bureau, l’API de compléments pour une utilisation dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="92ed6-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="92ed6-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="92ed6-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="92ed6-125">Membres</span><span class="sxs-lookup"><span data-stu-id="92ed6-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="92ed6-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="92ed6-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="92ed6-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="92ed6-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="92ed6-128">Type :</span><span class="sxs-lookup"><span data-stu-id="92ed6-128">Type:</span></span>

*   <span data-ttu-id="92ed6-129">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="92ed6-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="92ed6-130">Properties:</span></span>

|<span data-ttu-id="92ed6-131">Nom</span><span class="sxs-lookup"><span data-stu-id="92ed6-131">Name</span></span>| <span data-ttu-id="92ed6-132">Type</span><span class="sxs-lookup"><span data-stu-id="92ed6-132">Type</span></span>| <span data-ttu-id="92ed6-133">Description</span><span class="sxs-lookup"><span data-stu-id="92ed6-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="92ed6-134">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-134">String</span></span>|<span data-ttu-id="92ed6-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="92ed6-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="92ed6-136">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-136">String</span></span>|<span data-ttu-id="92ed6-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="92ed6-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92ed6-138">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92ed6-138">Requirements</span></span>

|<span data-ttu-id="92ed6-139">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92ed6-139">Requirement</span></span>| <span data-ttu-id="92ed6-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="92ed6-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="92ed6-141">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92ed6-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92ed6-142">1.0</span><span class="sxs-lookup"><span data-stu-id="92ed6-142">1.0</span></span>|
|[<span data-ttu-id="92ed6-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92ed6-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="92ed6-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92ed6-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="92ed6-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="92ed6-145">CoercionType :String</span></span>

<span data-ttu-id="92ed6-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="92ed6-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="92ed6-147">Type :</span><span class="sxs-lookup"><span data-stu-id="92ed6-147">Type:</span></span>

*   <span data-ttu-id="92ed6-148">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="92ed6-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="92ed6-149">Properties:</span></span>

|<span data-ttu-id="92ed6-150">Nom</span><span class="sxs-lookup"><span data-stu-id="92ed6-150">Name</span></span>| <span data-ttu-id="92ed6-151">Type</span><span class="sxs-lookup"><span data-stu-id="92ed6-151">Type</span></span>| <span data-ttu-id="92ed6-152">Description</span><span class="sxs-lookup"><span data-stu-id="92ed6-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="92ed6-153">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-153">String</span></span>|<span data-ttu-id="92ed6-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="92ed6-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="92ed6-155">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-155">String</span></span>|<span data-ttu-id="92ed6-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="92ed6-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92ed6-157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92ed6-157">Requirements</span></span>

|<span data-ttu-id="92ed6-158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92ed6-158">Requirement</span></span>| <span data-ttu-id="92ed6-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="92ed6-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="92ed6-160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92ed6-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92ed6-161">1.0</span><span class="sxs-lookup"><span data-stu-id="92ed6-161">1.0</span></span>|
|[<span data-ttu-id="92ed6-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92ed6-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="92ed6-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92ed6-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="92ed6-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="92ed6-164">EventType :String</span></span>

<span data-ttu-id="92ed6-165">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="92ed6-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="92ed6-166">Type :</span><span class="sxs-lookup"><span data-stu-id="92ed6-166">Type:</span></span>

*   <span data-ttu-id="92ed6-167">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="92ed6-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="92ed6-168">Properties:</span></span>

| <span data-ttu-id="92ed6-169">Nom</span><span class="sxs-lookup"><span data-stu-id="92ed6-169">Name</span></span> | <span data-ttu-id="92ed6-170">Type</span><span class="sxs-lookup"><span data-stu-id="92ed6-170">Type</span></span> | <span data-ttu-id="92ed6-171">Description</span><span class="sxs-lookup"><span data-stu-id="92ed6-171">Description</span></span> | <span data-ttu-id="92ed6-172">Ensemble de la configuration minimale requise</span><span class="sxs-lookup"><span data-stu-id="92ed6-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="92ed6-173">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-173">String</span></span> | <span data-ttu-id="92ed6-174">La date ou l’heure du rendez-vous sélectionné ou de série a changé.</span><span class="sxs-lookup"><span data-stu-id="92ed6-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="92ed6-175">1.7</span><span class="sxs-lookup"><span data-stu-id="92ed6-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="92ed6-176">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-176">String</span></span> | <span data-ttu-id="92ed6-177">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="92ed6-177">The selected item has changed.</span></span> | <span data-ttu-id="92ed6-178">1,5</span><span class="sxs-lookup"><span data-stu-id="92ed6-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="92ed6-179">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-179">String</span></span> | <span data-ttu-id="92ed6-180">La liste des destinataires de l’emplacement d’élément ou d’un rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="92ed6-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="92ed6-181">1.7</span><span class="sxs-lookup"><span data-stu-id="92ed6-181">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="92ed6-182">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-182">String</span></span> | <span data-ttu-id="92ed6-183">La périodicité de la série sélectionnée a changé.</span><span class="sxs-lookup"><span data-stu-id="92ed6-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="92ed6-184">1.7</span><span class="sxs-lookup"><span data-stu-id="92ed6-184">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="92ed6-185">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92ed6-185">Requirements</span></span>

|<span data-ttu-id="92ed6-186">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92ed6-186">Requirement</span></span>| <span data-ttu-id="92ed6-187">Valeur</span><span class="sxs-lookup"><span data-stu-id="92ed6-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="92ed6-188">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92ed6-188">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92ed6-189">1,5</span><span class="sxs-lookup"><span data-stu-id="92ed6-189">1.5</span></span> |
|[<span data-ttu-id="92ed6-190">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92ed6-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="92ed6-191">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92ed6-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="92ed6-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="92ed6-192">SourceProperty :String</span></span>

<span data-ttu-id="92ed6-193">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="92ed6-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="92ed6-194">Type :</span><span class="sxs-lookup"><span data-stu-id="92ed6-194">Type:</span></span>

*   <span data-ttu-id="92ed6-195">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="92ed6-196">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="92ed6-196">Properties:</span></span>

|<span data-ttu-id="92ed6-197">Nom</span><span class="sxs-lookup"><span data-stu-id="92ed6-197">Name</span></span>| <span data-ttu-id="92ed6-198">Type</span><span class="sxs-lookup"><span data-stu-id="92ed6-198">Type</span></span>| <span data-ttu-id="92ed6-199">Description</span><span class="sxs-lookup"><span data-stu-id="92ed6-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="92ed6-200">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-200">String</span></span>|<span data-ttu-id="92ed6-201">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="92ed6-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="92ed6-202">String</span><span class="sxs-lookup"><span data-stu-id="92ed6-202">String</span></span>|<span data-ttu-id="92ed6-203">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="92ed6-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92ed6-204">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="92ed6-204">Requirements</span></span>

|<span data-ttu-id="92ed6-205">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="92ed6-205">Requirement</span></span>| <span data-ttu-id="92ed6-206">Valeur</span><span class="sxs-lookup"><span data-stu-id="92ed6-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="92ed6-207">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="92ed6-207">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92ed6-208">1.0</span><span class="sxs-lookup"><span data-stu-id="92ed6-208">1.0</span></span>|
|[<span data-ttu-id="92ed6-209">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="92ed6-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="92ed6-210">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="92ed6-210">Compose or read</span></span>|