 

# <a name="office"></a><span data-ttu-id="22c1d-101">Bureau</span><span class="sxs-lookup"><span data-stu-id="22c1d-101">Office</span></span>

<span data-ttu-id="22c1d-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="22c1d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="22c1d-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="22c1d-104">Requirements</span></span>

|<span data-ttu-id="22c1d-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="22c1d-105">Requirement</span></span>| <span data-ttu-id="22c1d-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="22c1d-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="22c1d-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="22c1d-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22c1d-108">1.0</span><span class="sxs-lookup"><span data-stu-id="22c1d-108">1.0</span></span>|
|[<span data-ttu-id="22c1d-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="22c1d-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="22c1d-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="22c1d-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="22c1d-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="22c1d-111">Members and methods</span></span>

| <span data-ttu-id="22c1d-112">Membre</span><span class="sxs-lookup"><span data-stu-id="22c1d-112">Member</span></span> | <span data-ttu-id="22c1d-113">Type</span><span class="sxs-lookup"><span data-stu-id="22c1d-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="22c1d-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="22c1d-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="22c1d-115">Membre</span><span class="sxs-lookup"><span data-stu-id="22c1d-115">Member</span></span> |
| [<span data-ttu-id="22c1d-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="22c1d-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="22c1d-117">Membre</span><span class="sxs-lookup"><span data-stu-id="22c1d-117">Member</span></span> |
| [<span data-ttu-id="22c1d-118">EventType</span><span class="sxs-lookup"><span data-stu-id="22c1d-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="22c1d-119">Membre</span><span class="sxs-lookup"><span data-stu-id="22c1d-119">Member</span></span> |
| [<span data-ttu-id="22c1d-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="22c1d-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="22c1d-121">Membre</span><span class="sxs-lookup"><span data-stu-id="22c1d-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="22c1d-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="22c1d-122">Namespaces</span></span>

<span data-ttu-id="22c1d-123">[contexte](office.context.md): fournit des interfaces partagées d’espace de noms de contexte au bureau, l’API de compléments pour une utilisation dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="22c1d-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="22c1d-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="22c1d-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="22c1d-125">Membres</span><span class="sxs-lookup"><span data-stu-id="22c1d-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="22c1d-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="22c1d-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="22c1d-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="22c1d-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="22c1d-128">Type :</span><span class="sxs-lookup"><span data-stu-id="22c1d-128">Type:</span></span>

*   <span data-ttu-id="22c1d-129">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="22c1d-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="22c1d-130">Properties:</span></span>

|<span data-ttu-id="22c1d-131">Nom</span><span class="sxs-lookup"><span data-stu-id="22c1d-131">Name</span></span>| <span data-ttu-id="22c1d-132">Type</span><span class="sxs-lookup"><span data-stu-id="22c1d-132">Type</span></span>| <span data-ttu-id="22c1d-133">Description</span><span class="sxs-lookup"><span data-stu-id="22c1d-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="22c1d-134">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-134">String</span></span>|<span data-ttu-id="22c1d-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="22c1d-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="22c1d-136">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-136">String</span></span>|<span data-ttu-id="22c1d-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="22c1d-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22c1d-138">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="22c1d-138">Requirements</span></span>

|<span data-ttu-id="22c1d-139">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="22c1d-139">Requirement</span></span>| <span data-ttu-id="22c1d-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="22c1d-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="22c1d-141">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="22c1d-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22c1d-142">1.0</span><span class="sxs-lookup"><span data-stu-id="22c1d-142">1.0</span></span>|
|[<span data-ttu-id="22c1d-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="22c1d-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="22c1d-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="22c1d-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="22c1d-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="22c1d-145">CoercionType :String</span></span>

<span data-ttu-id="22c1d-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="22c1d-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="22c1d-147">Type :</span><span class="sxs-lookup"><span data-stu-id="22c1d-147">Type:</span></span>

*   <span data-ttu-id="22c1d-148">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="22c1d-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="22c1d-149">Properties:</span></span>

|<span data-ttu-id="22c1d-150">Nom</span><span class="sxs-lookup"><span data-stu-id="22c1d-150">Name</span></span>| <span data-ttu-id="22c1d-151">Type</span><span class="sxs-lookup"><span data-stu-id="22c1d-151">Type</span></span>| <span data-ttu-id="22c1d-152">Description</span><span class="sxs-lookup"><span data-stu-id="22c1d-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="22c1d-153">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-153">String</span></span>|<span data-ttu-id="22c1d-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="22c1d-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="22c1d-155">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-155">String</span></span>|<span data-ttu-id="22c1d-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="22c1d-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22c1d-157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="22c1d-157">Requirements</span></span>

|<span data-ttu-id="22c1d-158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="22c1d-158">Requirement</span></span>| <span data-ttu-id="22c1d-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="22c1d-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="22c1d-160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="22c1d-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22c1d-161">1.0</span><span class="sxs-lookup"><span data-stu-id="22c1d-161">1.0</span></span>|
|[<span data-ttu-id="22c1d-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="22c1d-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="22c1d-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="22c1d-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="22c1d-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="22c1d-164">EventType :String</span></span>

<span data-ttu-id="22c1d-165">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="22c1d-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="22c1d-166">Type :</span><span class="sxs-lookup"><span data-stu-id="22c1d-166">Type:</span></span>

*   <span data-ttu-id="22c1d-167">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="22c1d-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="22c1d-168">Properties:</span></span>

| <span data-ttu-id="22c1d-169">Nom</span><span class="sxs-lookup"><span data-stu-id="22c1d-169">Name</span></span> | <span data-ttu-id="22c1d-170">Type</span><span class="sxs-lookup"><span data-stu-id="22c1d-170">Type</span></span> | <span data-ttu-id="22c1d-171">Description</span><span class="sxs-lookup"><span data-stu-id="22c1d-171">Description</span></span> | <span data-ttu-id="22c1d-172">Ensemble de la configuration minimale requise</span><span class="sxs-lookup"><span data-stu-id="22c1d-172">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="22c1d-173">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-173">String</span></span> | <span data-ttu-id="22c1d-174">Le rendez-vous date ou l’heure de la série sélectionnée a changé.</span><span class="sxs-lookup"><span data-stu-id="22c1d-174">The appointment date or time of the selected series has changed.</span></span> | <span data-ttu-id="22c1d-175">Preview</span><span class="sxs-lookup"><span data-stu-id="22c1d-175">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="22c1d-176">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-176">String</span></span> | <span data-ttu-id="22c1d-177">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="22c1d-177">The selected item has changed.</span></span> | <span data-ttu-id="22c1d-178">1,5</span><span class="sxs-lookup"><span data-stu-id="22c1d-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="22c1d-179">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-179">String</span></span> | <span data-ttu-id="22c1d-180">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="22c1d-180">The selected item has changed.</span></span> | <span data-ttu-id="22c1d-181">Preview</span><span class="sxs-lookup"><span data-stu-id="22c1d-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="22c1d-182">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-182">String</span></span> | <span data-ttu-id="22c1d-183">La liste des destinataires de l’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="22c1d-183">The recipient list of the selected item has changed.</span></span> | <span data-ttu-id="22c1d-184">Preview</span><span class="sxs-lookup"><span data-stu-id="22c1d-184">Preview</span></span> |
|`RecurrencePatternChanged`| <span data-ttu-id="22c1d-185">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-185">String</span></span> | <span data-ttu-id="22c1d-186">La périodicité de la série sélectionnée a changé.</span><span class="sxs-lookup"><span data-stu-id="22c1d-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="22c1d-187">Preview</span><span class="sxs-lookup"><span data-stu-id="22c1d-187">Preview</span></span> |

##### <a name="requirements"></a><span data-ttu-id="22c1d-188">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="22c1d-188">Requirements</span></span>

|<span data-ttu-id="22c1d-189">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="22c1d-189">Requirement</span></span>| <span data-ttu-id="22c1d-190">Valeur</span><span class="sxs-lookup"><span data-stu-id="22c1d-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="22c1d-191">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="22c1d-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22c1d-192">1,5</span><span class="sxs-lookup"><span data-stu-id="22c1d-192">1.5</span></span> |
|[<span data-ttu-id="22c1d-193">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="22c1d-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="22c1d-194">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="22c1d-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="22c1d-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="22c1d-195">SourceProperty :String</span></span>

<span data-ttu-id="22c1d-196">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="22c1d-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="22c1d-197">Type :</span><span class="sxs-lookup"><span data-stu-id="22c1d-197">Type:</span></span>

*   <span data-ttu-id="22c1d-198">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="22c1d-199">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="22c1d-199">Properties:</span></span>

|<span data-ttu-id="22c1d-200">Nom</span><span class="sxs-lookup"><span data-stu-id="22c1d-200">Name</span></span>| <span data-ttu-id="22c1d-201">Type</span><span class="sxs-lookup"><span data-stu-id="22c1d-201">Type</span></span>| <span data-ttu-id="22c1d-202">Description</span><span class="sxs-lookup"><span data-stu-id="22c1d-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="22c1d-203">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-203">String</span></span>|<span data-ttu-id="22c1d-204">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="22c1d-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="22c1d-205">String</span><span class="sxs-lookup"><span data-stu-id="22c1d-205">String</span></span>|<span data-ttu-id="22c1d-206">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="22c1d-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22c1d-207">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="22c1d-207">Requirements</span></span>

|<span data-ttu-id="22c1d-208">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="22c1d-208">Requirement</span></span>| <span data-ttu-id="22c1d-209">Valeur</span><span class="sxs-lookup"><span data-stu-id="22c1d-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="22c1d-210">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="22c1d-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22c1d-211">1.0</span><span class="sxs-lookup"><span data-stu-id="22c1d-211">1.0</span></span>|
|[<span data-ttu-id="22c1d-212">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="22c1d-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="22c1d-213">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="22c1d-213">Compose or read</span></span>|