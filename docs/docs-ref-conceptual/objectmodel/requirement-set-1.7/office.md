 

# <a name="office"></a><span data-ttu-id="4769c-101">Bureau</span><span class="sxs-lookup"><span data-stu-id="4769c-101">Office</span></span>

<span data-ttu-id="4769c-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="4769c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4769c-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4769c-104">Requirements</span></span>

|<span data-ttu-id="4769c-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4769c-105">Requirement</span></span>| <span data-ttu-id="4769c-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="4769c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4769c-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4769c-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4769c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4769c-108">1.0</span></span>|
|[<span data-ttu-id="4769c-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4769c-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4769c-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4769c-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4769c-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4769c-111">Members and methods</span></span>

| <span data-ttu-id="4769c-112">Membre</span><span class="sxs-lookup"><span data-stu-id="4769c-112">Member</span></span> | <span data-ttu-id="4769c-113">Type</span><span class="sxs-lookup"><span data-stu-id="4769c-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4769c-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4769c-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4769c-115">Membre</span><span class="sxs-lookup"><span data-stu-id="4769c-115">Member</span></span> |
| [<span data-ttu-id="4769c-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4769c-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4769c-117">Membre</span><span class="sxs-lookup"><span data-stu-id="4769c-117">Member</span></span> |
| [<span data-ttu-id="4769c-118">EventType</span><span class="sxs-lookup"><span data-stu-id="4769c-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4769c-119">Membre</span><span class="sxs-lookup"><span data-stu-id="4769c-119">Member</span></span> |
| [<span data-ttu-id="4769c-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4769c-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4769c-121">Membre</span><span class="sxs-lookup"><span data-stu-id="4769c-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4769c-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="4769c-122">Namespaces</span></span>

<span data-ttu-id="4769c-123">[contexte](office.context.md): fournit des interfaces partagées d’espace de noms de contexte au bureau, l’API de compléments pour une utilisation dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4769c-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="4769c-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="4769c-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="4769c-125">Membres</span><span class="sxs-lookup"><span data-stu-id="4769c-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="4769c-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="4769c-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="4769c-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4769c-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4769c-128">Type :</span><span class="sxs-lookup"><span data-stu-id="4769c-128">Type:</span></span>

*   <span data-ttu-id="4769c-129">String</span><span class="sxs-lookup"><span data-stu-id="4769c-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4769c-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4769c-130">Properties:</span></span>

|<span data-ttu-id="4769c-131">Nom</span><span class="sxs-lookup"><span data-stu-id="4769c-131">Name</span></span>| <span data-ttu-id="4769c-132">Type</span><span class="sxs-lookup"><span data-stu-id="4769c-132">Type</span></span>| <span data-ttu-id="4769c-133">Description</span><span class="sxs-lookup"><span data-stu-id="4769c-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4769c-134">String</span><span class="sxs-lookup"><span data-stu-id="4769c-134">String</span></span>|<span data-ttu-id="4769c-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="4769c-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4769c-136">String</span><span class="sxs-lookup"><span data-stu-id="4769c-136">String</span></span>|<span data-ttu-id="4769c-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="4769c-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4769c-138">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4769c-138">Requirements</span></span>

|<span data-ttu-id="4769c-139">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4769c-139">Requirement</span></span>| <span data-ttu-id="4769c-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="4769c-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="4769c-141">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4769c-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4769c-142">1.0</span><span class="sxs-lookup"><span data-stu-id="4769c-142">1.0</span></span>|
|[<span data-ttu-id="4769c-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4769c-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4769c-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4769c-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="4769c-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="4769c-145">CoercionType :String</span></span>

<span data-ttu-id="4769c-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4769c-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4769c-147">Type :</span><span class="sxs-lookup"><span data-stu-id="4769c-147">Type:</span></span>

*   <span data-ttu-id="4769c-148">String</span><span class="sxs-lookup"><span data-stu-id="4769c-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4769c-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4769c-149">Properties:</span></span>

|<span data-ttu-id="4769c-150">Nom</span><span class="sxs-lookup"><span data-stu-id="4769c-150">Name</span></span>| <span data-ttu-id="4769c-151">Type</span><span class="sxs-lookup"><span data-stu-id="4769c-151">Type</span></span>| <span data-ttu-id="4769c-152">Description</span><span class="sxs-lookup"><span data-stu-id="4769c-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4769c-153">String</span><span class="sxs-lookup"><span data-stu-id="4769c-153">String</span></span>|<span data-ttu-id="4769c-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="4769c-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4769c-155">String</span><span class="sxs-lookup"><span data-stu-id="4769c-155">String</span></span>|<span data-ttu-id="4769c-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="4769c-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4769c-157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4769c-157">Requirements</span></span>

|<span data-ttu-id="4769c-158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4769c-158">Requirement</span></span>| <span data-ttu-id="4769c-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="4769c-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="4769c-160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4769c-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4769c-161">1.0</span><span class="sxs-lookup"><span data-stu-id="4769c-161">1.0</span></span>|
|[<span data-ttu-id="4769c-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4769c-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4769c-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4769c-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="4769c-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="4769c-164">EventType :String</span></span>

<span data-ttu-id="4769c-165">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="4769c-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4769c-166">Type :</span><span class="sxs-lookup"><span data-stu-id="4769c-166">Type:</span></span>

*   <span data-ttu-id="4769c-167">String</span><span class="sxs-lookup"><span data-stu-id="4769c-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4769c-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4769c-168">Properties:</span></span>

| <span data-ttu-id="4769c-169">Nom</span><span class="sxs-lookup"><span data-stu-id="4769c-169">Name</span></span> | <span data-ttu-id="4769c-170">Type</span><span class="sxs-lookup"><span data-stu-id="4769c-170">Type</span></span> | <span data-ttu-id="4769c-171">Description</span><span class="sxs-lookup"><span data-stu-id="4769c-171">Description</span></span> | <span data-ttu-id="4769c-172">Ensemble de la configuration minimale requise</span><span class="sxs-lookup"><span data-stu-id="4769c-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="4769c-173">String</span><span class="sxs-lookup"><span data-stu-id="4769c-173">String</span></span> | <span data-ttu-id="4769c-174">La date ou l’heure du rendez-vous sélectionné ou de série a changé.</span><span class="sxs-lookup"><span data-stu-id="4769c-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="4769c-175">1.7</span><span class="sxs-lookup"><span data-stu-id="4769c-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="4769c-176">String</span><span class="sxs-lookup"><span data-stu-id="4769c-176">String</span></span> | <span data-ttu-id="4769c-177">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="4769c-177">The selected item has changed.</span></span> | <span data-ttu-id="4769c-178">1,5</span><span class="sxs-lookup"><span data-stu-id="4769c-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="4769c-179">String</span><span class="sxs-lookup"><span data-stu-id="4769c-179">String</span></span> | <span data-ttu-id="4769c-180">La liste des destinataires de l’emplacement d’élément ou d’un rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="4769c-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="4769c-181">1.7</span><span class="sxs-lookup"><span data-stu-id="4769c-181">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="4769c-182">String</span><span class="sxs-lookup"><span data-stu-id="4769c-182">String</span></span> | <span data-ttu-id="4769c-183">La périodicité de la série sélectionnée a changé.</span><span class="sxs-lookup"><span data-stu-id="4769c-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="4769c-184">1.7</span><span class="sxs-lookup"><span data-stu-id="4769c-184">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4769c-185">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4769c-185">Requirements</span></span>

|<span data-ttu-id="4769c-186">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4769c-186">Requirement</span></span>| <span data-ttu-id="4769c-187">Valeur</span><span class="sxs-lookup"><span data-stu-id="4769c-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="4769c-188">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4769c-188">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4769c-189">1,5</span><span class="sxs-lookup"><span data-stu-id="4769c-189">1.5</span></span> |
|[<span data-ttu-id="4769c-190">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4769c-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4769c-191">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4769c-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="4769c-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="4769c-192">SourceProperty :String</span></span>

<span data-ttu-id="4769c-193">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4769c-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4769c-194">Type :</span><span class="sxs-lookup"><span data-stu-id="4769c-194">Type:</span></span>

*   <span data-ttu-id="4769c-195">String</span><span class="sxs-lookup"><span data-stu-id="4769c-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4769c-196">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4769c-196">Properties:</span></span>

|<span data-ttu-id="4769c-197">Nom</span><span class="sxs-lookup"><span data-stu-id="4769c-197">Name</span></span>| <span data-ttu-id="4769c-198">Type</span><span class="sxs-lookup"><span data-stu-id="4769c-198">Type</span></span>| <span data-ttu-id="4769c-199">Description</span><span class="sxs-lookup"><span data-stu-id="4769c-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4769c-200">String</span><span class="sxs-lookup"><span data-stu-id="4769c-200">String</span></span>|<span data-ttu-id="4769c-201">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="4769c-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4769c-202">String</span><span class="sxs-lookup"><span data-stu-id="4769c-202">String</span></span>|<span data-ttu-id="4769c-203">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="4769c-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4769c-204">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4769c-204">Requirements</span></span>

|<span data-ttu-id="4769c-205">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4769c-205">Requirement</span></span>| <span data-ttu-id="4769c-206">Valeur</span><span class="sxs-lookup"><span data-stu-id="4769c-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="4769c-207">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4769c-207">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4769c-208">1.0</span><span class="sxs-lookup"><span data-stu-id="4769c-208">1.0</span></span>|
|[<span data-ttu-id="4769c-209">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4769c-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4769c-210">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4769c-210">Compose or read</span></span>|