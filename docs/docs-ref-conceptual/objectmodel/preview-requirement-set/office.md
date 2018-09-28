 

# <a name="office"></a><span data-ttu-id="38194-101">Bureau</span><span class="sxs-lookup"><span data-stu-id="38194-101">Office</span></span>

<span data-ttu-id="38194-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="38194-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="38194-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="38194-104">Requirements</span></span>

|<span data-ttu-id="38194-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="38194-105">Requirement</span></span>| <span data-ttu-id="38194-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="38194-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="38194-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="38194-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38194-108">1.0</span><span class="sxs-lookup"><span data-stu-id="38194-108">1.0</span></span>|
|[<span data-ttu-id="38194-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="38194-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38194-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="38194-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="38194-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="38194-111">Members and methods</span></span>

| <span data-ttu-id="38194-112">Membre</span><span class="sxs-lookup"><span data-stu-id="38194-112">Member</span></span> | <span data-ttu-id="38194-113">Type</span><span class="sxs-lookup"><span data-stu-id="38194-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="38194-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="38194-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="38194-115">Membre</span><span class="sxs-lookup"><span data-stu-id="38194-115">Member</span></span> |
| [<span data-ttu-id="38194-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="38194-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="38194-117">Membre</span><span class="sxs-lookup"><span data-stu-id="38194-117">Member</span></span> |
| [<span data-ttu-id="38194-118">EventType</span><span class="sxs-lookup"><span data-stu-id="38194-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="38194-119">Membre</span><span class="sxs-lookup"><span data-stu-id="38194-119">Member</span></span> |
| [<span data-ttu-id="38194-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="38194-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="38194-121">Membre</span><span class="sxs-lookup"><span data-stu-id="38194-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="38194-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="38194-122">Namespaces</span></span>

<span data-ttu-id="38194-123">[contexte](office.context.md): fournit des interfaces partagées d’espace de noms de contexte au bureau, l’API de compléments pour une utilisation dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="38194-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="38194-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="38194-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="38194-125">Membres</span><span class="sxs-lookup"><span data-stu-id="38194-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="38194-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="38194-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="38194-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="38194-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="38194-128">Type :</span><span class="sxs-lookup"><span data-stu-id="38194-128">Type:</span></span>

*   <span data-ttu-id="38194-129">String</span><span class="sxs-lookup"><span data-stu-id="38194-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="38194-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="38194-130">Properties:</span></span>

|<span data-ttu-id="38194-131">Nom</span><span class="sxs-lookup"><span data-stu-id="38194-131">Name</span></span>| <span data-ttu-id="38194-132">Type</span><span class="sxs-lookup"><span data-stu-id="38194-132">Type</span></span>| <span data-ttu-id="38194-133">Description</span><span class="sxs-lookup"><span data-stu-id="38194-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="38194-134">String</span><span class="sxs-lookup"><span data-stu-id="38194-134">String</span></span>|<span data-ttu-id="38194-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="38194-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="38194-136">String</span><span class="sxs-lookup"><span data-stu-id="38194-136">String</span></span>|<span data-ttu-id="38194-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="38194-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38194-138">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="38194-138">Requirements</span></span>

|<span data-ttu-id="38194-139">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="38194-139">Requirement</span></span>| <span data-ttu-id="38194-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="38194-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="38194-141">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="38194-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38194-142">1.0</span><span class="sxs-lookup"><span data-stu-id="38194-142">1.0</span></span>|
|[<span data-ttu-id="38194-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="38194-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38194-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="38194-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="38194-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="38194-145">CoercionType :String</span></span>

<span data-ttu-id="38194-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="38194-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="38194-147">Type :</span><span class="sxs-lookup"><span data-stu-id="38194-147">Type:</span></span>

*   <span data-ttu-id="38194-148">String</span><span class="sxs-lookup"><span data-stu-id="38194-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="38194-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="38194-149">Properties:</span></span>

|<span data-ttu-id="38194-150">Nom</span><span class="sxs-lookup"><span data-stu-id="38194-150">Name</span></span>| <span data-ttu-id="38194-151">Type</span><span class="sxs-lookup"><span data-stu-id="38194-151">Type</span></span>| <span data-ttu-id="38194-152">Description</span><span class="sxs-lookup"><span data-stu-id="38194-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="38194-153">String</span><span class="sxs-lookup"><span data-stu-id="38194-153">String</span></span>|<span data-ttu-id="38194-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="38194-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="38194-155">String</span><span class="sxs-lookup"><span data-stu-id="38194-155">String</span></span>|<span data-ttu-id="38194-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="38194-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38194-157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="38194-157">Requirements</span></span>

|<span data-ttu-id="38194-158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="38194-158">Requirement</span></span>| <span data-ttu-id="38194-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="38194-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="38194-160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="38194-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38194-161">1.0</span><span class="sxs-lookup"><span data-stu-id="38194-161">1.0</span></span>|
|[<span data-ttu-id="38194-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="38194-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38194-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="38194-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="38194-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="38194-164">EventType :String</span></span>

<span data-ttu-id="38194-165">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="38194-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="38194-166">Type :</span><span class="sxs-lookup"><span data-stu-id="38194-166">Type:</span></span>

*   <span data-ttu-id="38194-167">String</span><span class="sxs-lookup"><span data-stu-id="38194-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="38194-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="38194-168">Properties:</span></span>

| <span data-ttu-id="38194-169">Nom</span><span class="sxs-lookup"><span data-stu-id="38194-169">Name</span></span> | <span data-ttu-id="38194-170">Type</span><span class="sxs-lookup"><span data-stu-id="38194-170">Type</span></span> | <span data-ttu-id="38194-171">Description</span><span class="sxs-lookup"><span data-stu-id="38194-171">Description</span></span> | <span data-ttu-id="38194-172">Ensemble de la configuration minimale requise</span><span class="sxs-lookup"><span data-stu-id="38194-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="38194-173">String</span><span class="sxs-lookup"><span data-stu-id="38194-173">String</span></span> | <span data-ttu-id="38194-174">La date ou l’heure du rendez-vous sélectionné ou de série a changé.</span><span class="sxs-lookup"><span data-stu-id="38194-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="38194-175">1.7</span><span class="sxs-lookup"><span data-stu-id="38194-175">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="38194-176">String</span><span class="sxs-lookup"><span data-stu-id="38194-176">String</span></span> | <span data-ttu-id="38194-177">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="38194-177">The selected item has changed.</span></span> | <span data-ttu-id="38194-178">1,5</span><span class="sxs-lookup"><span data-stu-id="38194-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="38194-179">String</span><span class="sxs-lookup"><span data-stu-id="38194-179">String</span></span> | <span data-ttu-id="38194-180">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="38194-180">The selected item has changed.</span></span> | <span data-ttu-id="38194-181">Preview</span><span class="sxs-lookup"><span data-stu-id="38194-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="38194-182">String</span><span class="sxs-lookup"><span data-stu-id="38194-182">String</span></span> | <span data-ttu-id="38194-183">La liste des destinataires de l’emplacement d’élément ou d’un rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="38194-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="38194-184">1.7</span><span class="sxs-lookup"><span data-stu-id="38194-184">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="38194-185">String</span><span class="sxs-lookup"><span data-stu-id="38194-185">String</span></span> | <span data-ttu-id="38194-186">La périodicité de la série sélectionnée a changé.</span><span class="sxs-lookup"><span data-stu-id="38194-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="38194-187">1.7</span><span class="sxs-lookup"><span data-stu-id="38194-187">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="38194-188">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="38194-188">Requirements</span></span>

|<span data-ttu-id="38194-189">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="38194-189">Requirement</span></span>| <span data-ttu-id="38194-190">Valeur</span><span class="sxs-lookup"><span data-stu-id="38194-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="38194-191">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="38194-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38194-192">1,5</span><span class="sxs-lookup"><span data-stu-id="38194-192">1.5</span></span> |
|[<span data-ttu-id="38194-193">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="38194-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38194-194">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="38194-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="38194-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="38194-195">SourceProperty :String</span></span>

<span data-ttu-id="38194-196">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="38194-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="38194-197">Type :</span><span class="sxs-lookup"><span data-stu-id="38194-197">Type:</span></span>

*   <span data-ttu-id="38194-198">String</span><span class="sxs-lookup"><span data-stu-id="38194-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="38194-199">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="38194-199">Properties:</span></span>

|<span data-ttu-id="38194-200">Nom</span><span class="sxs-lookup"><span data-stu-id="38194-200">Name</span></span>| <span data-ttu-id="38194-201">Type</span><span class="sxs-lookup"><span data-stu-id="38194-201">Type</span></span>| <span data-ttu-id="38194-202">Description</span><span class="sxs-lookup"><span data-stu-id="38194-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="38194-203">String</span><span class="sxs-lookup"><span data-stu-id="38194-203">String</span></span>|<span data-ttu-id="38194-204">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="38194-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="38194-205">String</span><span class="sxs-lookup"><span data-stu-id="38194-205">String</span></span>|<span data-ttu-id="38194-206">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="38194-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38194-207">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="38194-207">Requirements</span></span>

|<span data-ttu-id="38194-208">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="38194-208">Requirement</span></span>| <span data-ttu-id="38194-209">Valeur</span><span class="sxs-lookup"><span data-stu-id="38194-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="38194-210">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="38194-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38194-211">1.0</span><span class="sxs-lookup"><span data-stu-id="38194-211">1.0</span></span>|
|[<span data-ttu-id="38194-212">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="38194-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38194-213">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="38194-213">Compose or read</span></span>|