# <a name="office"></a><span data-ttu-id="486af-101">Bureau</span><span class="sxs-lookup"><span data-stu-id="486af-101">Office</span></span>

<span data-ttu-id="486af-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="486af-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="486af-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="486af-104">Requirements</span></span>

|<span data-ttu-id="486af-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="486af-105">Requirement</span></span>| <span data-ttu-id="486af-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="486af-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="486af-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="486af-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="486af-108">1.0</span><span class="sxs-lookup"><span data-stu-id="486af-108">1.0</span></span>|
|[<span data-ttu-id="486af-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="486af-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="486af-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="486af-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="486af-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="486af-111">Members and methods</span></span>

| <span data-ttu-id="486af-112">Membre</span><span class="sxs-lookup"><span data-stu-id="486af-112">Member</span></span> | <span data-ttu-id="486af-113">Type</span><span class="sxs-lookup"><span data-stu-id="486af-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="486af-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="486af-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="486af-115">Membre</span><span class="sxs-lookup"><span data-stu-id="486af-115">Member</span></span> |
| [<span data-ttu-id="486af-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="486af-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="486af-117">Membre</span><span class="sxs-lookup"><span data-stu-id="486af-117">Member</span></span> |
| [<span data-ttu-id="486af-118">EventType</span><span class="sxs-lookup"><span data-stu-id="486af-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="486af-119">Membre</span><span class="sxs-lookup"><span data-stu-id="486af-119">Member</span></span> |
| [<span data-ttu-id="486af-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="486af-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="486af-121">Membre</span><span class="sxs-lookup"><span data-stu-id="486af-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="486af-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="486af-122">Namespaces</span></span>

<span data-ttu-id="486af-123">[contexte](office.context.md): fournit des interfaces partagées d’espace de noms de contexte au bureau, l’API de compléments pour une utilisation dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="486af-123">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="486af-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="486af-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="486af-125">Membres</span><span class="sxs-lookup"><span data-stu-id="486af-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="486af-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="486af-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="486af-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="486af-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="486af-128">Type :</span><span class="sxs-lookup"><span data-stu-id="486af-128">Type:</span></span>

*   <span data-ttu-id="486af-129">String</span><span class="sxs-lookup"><span data-stu-id="486af-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="486af-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="486af-130">Properties:</span></span>

|<span data-ttu-id="486af-131">Nom</span><span class="sxs-lookup"><span data-stu-id="486af-131">Name</span></span>| <span data-ttu-id="486af-132">Type</span><span class="sxs-lookup"><span data-stu-id="486af-132">Type</span></span>| <span data-ttu-id="486af-133">Description</span><span class="sxs-lookup"><span data-stu-id="486af-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="486af-134">String</span><span class="sxs-lookup"><span data-stu-id="486af-134">String</span></span>|<span data-ttu-id="486af-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="486af-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="486af-136">String</span><span class="sxs-lookup"><span data-stu-id="486af-136">String</span></span>|<span data-ttu-id="486af-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="486af-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="486af-138">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="486af-138">Requirements</span></span>

|<span data-ttu-id="486af-139">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="486af-139">Requirement</span></span>| <span data-ttu-id="486af-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="486af-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="486af-141">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="486af-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="486af-142">1.0</span><span class="sxs-lookup"><span data-stu-id="486af-142">1.0</span></span>|
|[<span data-ttu-id="486af-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="486af-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="486af-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="486af-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="486af-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="486af-145">CoercionType :String</span></span>

<span data-ttu-id="486af-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="486af-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="486af-147">Type :</span><span class="sxs-lookup"><span data-stu-id="486af-147">Type:</span></span>

*   <span data-ttu-id="486af-148">String</span><span class="sxs-lookup"><span data-stu-id="486af-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="486af-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="486af-149">Properties:</span></span>

|<span data-ttu-id="486af-150">Nom</span><span class="sxs-lookup"><span data-stu-id="486af-150">Name</span></span>| <span data-ttu-id="486af-151">Type</span><span class="sxs-lookup"><span data-stu-id="486af-151">Type</span></span>| <span data-ttu-id="486af-152">Description</span><span class="sxs-lookup"><span data-stu-id="486af-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="486af-153">String</span><span class="sxs-lookup"><span data-stu-id="486af-153">String</span></span>|<span data-ttu-id="486af-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="486af-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="486af-155">String</span><span class="sxs-lookup"><span data-stu-id="486af-155">String</span></span>|<span data-ttu-id="486af-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="486af-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="486af-157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="486af-157">Requirements</span></span>

|<span data-ttu-id="486af-158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="486af-158">Requirement</span></span>| <span data-ttu-id="486af-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="486af-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="486af-160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="486af-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="486af-161">1.0</span><span class="sxs-lookup"><span data-stu-id="486af-161">1.0</span></span>|
|[<span data-ttu-id="486af-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="486af-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="486af-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="486af-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="486af-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="486af-164">EventType :String</span></span>

<span data-ttu-id="486af-165">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="486af-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="486af-166">Type :</span><span class="sxs-lookup"><span data-stu-id="486af-166">Type:</span></span>

*   <span data-ttu-id="486af-167">String</span><span class="sxs-lookup"><span data-stu-id="486af-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="486af-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="486af-168">Properties:</span></span>

| <span data-ttu-id="486af-169">Nom</span><span class="sxs-lookup"><span data-stu-id="486af-169">Name</span></span> | <span data-ttu-id="486af-170">Type</span><span class="sxs-lookup"><span data-stu-id="486af-170">Type</span></span> | <span data-ttu-id="486af-171">Description</span><span class="sxs-lookup"><span data-stu-id="486af-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="486af-172">String</span><span class="sxs-lookup"><span data-stu-id="486af-172">String</span></span> | <span data-ttu-id="486af-173">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="486af-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="486af-174">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="486af-174">Requirements</span></span>

|<span data-ttu-id="486af-175">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="486af-175">Requirement</span></span>| <span data-ttu-id="486af-176">Valeur</span><span class="sxs-lookup"><span data-stu-id="486af-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="486af-177">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="486af-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="486af-178">1,5</span><span class="sxs-lookup"><span data-stu-id="486af-178">1.5</span></span> |
|[<span data-ttu-id="486af-179">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="486af-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="486af-180">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="486af-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="486af-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="486af-181">SourceProperty :String</span></span>

<span data-ttu-id="486af-182">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="486af-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="486af-183">Type :</span><span class="sxs-lookup"><span data-stu-id="486af-183">Type:</span></span>

*   <span data-ttu-id="486af-184">String</span><span class="sxs-lookup"><span data-stu-id="486af-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="486af-185">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="486af-185">Properties:</span></span>

|<span data-ttu-id="486af-186">Nom</span><span class="sxs-lookup"><span data-stu-id="486af-186">Name</span></span>| <span data-ttu-id="486af-187">Type</span><span class="sxs-lookup"><span data-stu-id="486af-187">Type</span></span>| <span data-ttu-id="486af-188">Description</span><span class="sxs-lookup"><span data-stu-id="486af-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="486af-189">String</span><span class="sxs-lookup"><span data-stu-id="486af-189">String</span></span>|<span data-ttu-id="486af-190">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="486af-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="486af-191">String</span><span class="sxs-lookup"><span data-stu-id="486af-191">String</span></span>|<span data-ttu-id="486af-192">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="486af-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="486af-193">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="486af-193">Requirements</span></span>

|<span data-ttu-id="486af-194">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="486af-194">Requirement</span></span>| <span data-ttu-id="486af-195">Valeur</span><span class="sxs-lookup"><span data-stu-id="486af-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="486af-196">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="486af-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="486af-197">1.0</span><span class="sxs-lookup"><span data-stu-id="486af-197">1.0</span></span>|
|[<span data-ttu-id="486af-198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="486af-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="486af-199">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="486af-199">Compose or read</span></span>|