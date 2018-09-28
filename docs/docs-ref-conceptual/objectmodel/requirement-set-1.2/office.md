 

# <a name="office"></a><span data-ttu-id="be26c-101">Bureau</span><span class="sxs-lookup"><span data-stu-id="be26c-101">Office</span></span>

<span data-ttu-id="be26c-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="be26c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26c-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be26c-104">Requirements</span></span>

|<span data-ttu-id="be26c-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be26c-105">Requirement</span></span>| <span data-ttu-id="be26c-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="be26c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26c-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be26c-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="be26c-108">1.0</span></span>|
|[<span data-ttu-id="be26c-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be26c-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26c-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="be26c-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="be26c-111">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="be26c-111">Namespaces</span></span>

<span data-ttu-id="be26c-112">[contexte](office.context.md): fournit des interfaces partagées d’espace de noms de contexte au bureau, l’API de compléments pour une utilisation dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="be26c-112">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="be26c-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="be26c-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="be26c-114">Membres</span><span class="sxs-lookup"><span data-stu-id="be26c-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="be26c-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="be26c-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="be26c-116">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="be26c-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="be26c-117">Type :</span><span class="sxs-lookup"><span data-stu-id="be26c-117">Type:</span></span>

*   <span data-ttu-id="be26c-118">String</span><span class="sxs-lookup"><span data-stu-id="be26c-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="be26c-119">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="be26c-119">Properties:</span></span>

|<span data-ttu-id="be26c-120">Nom</span><span class="sxs-lookup"><span data-stu-id="be26c-120">Name</span></span>| <span data-ttu-id="be26c-121">Type</span><span class="sxs-lookup"><span data-stu-id="be26c-121">Type</span></span>| <span data-ttu-id="be26c-122">Description</span><span class="sxs-lookup"><span data-stu-id="be26c-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="be26c-123">String</span><span class="sxs-lookup"><span data-stu-id="be26c-123">String</span></span>|<span data-ttu-id="be26c-124">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="be26c-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="be26c-125">String</span><span class="sxs-lookup"><span data-stu-id="be26c-125">String</span></span>|<span data-ttu-id="be26c-126">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="be26c-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be26c-127">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be26c-127">Requirements</span></span>

|<span data-ttu-id="be26c-128">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be26c-128">Requirement</span></span>| <span data-ttu-id="be26c-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="be26c-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26c-130">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be26c-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26c-131">1.0</span><span class="sxs-lookup"><span data-stu-id="be26c-131">1.0</span></span>|
|[<span data-ttu-id="be26c-132">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be26c-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26c-133">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="be26c-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="be26c-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="be26c-134">CoercionType :String</span></span>

<span data-ttu-id="be26c-135">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="be26c-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="be26c-136">Type :</span><span class="sxs-lookup"><span data-stu-id="be26c-136">Type:</span></span>

*   <span data-ttu-id="be26c-137">String</span><span class="sxs-lookup"><span data-stu-id="be26c-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="be26c-138">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="be26c-138">Properties:</span></span>

|<span data-ttu-id="be26c-139">Nom</span><span class="sxs-lookup"><span data-stu-id="be26c-139">Name</span></span>| <span data-ttu-id="be26c-140">Type</span><span class="sxs-lookup"><span data-stu-id="be26c-140">Type</span></span>| <span data-ttu-id="be26c-141">Description</span><span class="sxs-lookup"><span data-stu-id="be26c-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="be26c-142">String</span><span class="sxs-lookup"><span data-stu-id="be26c-142">String</span></span>|<span data-ttu-id="be26c-143">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="be26c-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="be26c-144">String</span><span class="sxs-lookup"><span data-stu-id="be26c-144">String</span></span>|<span data-ttu-id="be26c-145">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="be26c-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be26c-146">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be26c-146">Requirements</span></span>

|<span data-ttu-id="be26c-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be26c-147">Requirement</span></span>| <span data-ttu-id="be26c-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="be26c-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26c-149">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be26c-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26c-150">1.0</span><span class="sxs-lookup"><span data-stu-id="be26c-150">1.0</span></span>|
|[<span data-ttu-id="be26c-151">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be26c-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26c-152">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="be26c-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="be26c-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="be26c-153">SourceProperty :String</span></span>

<span data-ttu-id="be26c-154">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="be26c-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="be26c-155">Type :</span><span class="sxs-lookup"><span data-stu-id="be26c-155">Type:</span></span>

*   <span data-ttu-id="be26c-156">String</span><span class="sxs-lookup"><span data-stu-id="be26c-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="be26c-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="be26c-157">Properties:</span></span>

|<span data-ttu-id="be26c-158">Nom</span><span class="sxs-lookup"><span data-stu-id="be26c-158">Name</span></span>| <span data-ttu-id="be26c-159">Type</span><span class="sxs-lookup"><span data-stu-id="be26c-159">Type</span></span>| <span data-ttu-id="be26c-160">Description</span><span class="sxs-lookup"><span data-stu-id="be26c-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="be26c-161">String</span><span class="sxs-lookup"><span data-stu-id="be26c-161">String</span></span>|<span data-ttu-id="be26c-162">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="be26c-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="be26c-163">String</span><span class="sxs-lookup"><span data-stu-id="be26c-163">String</span></span>|<span data-ttu-id="be26c-164">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="be26c-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be26c-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be26c-165">Requirements</span></span>

|<span data-ttu-id="be26c-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be26c-166">Requirement</span></span>| <span data-ttu-id="be26c-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="be26c-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26c-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be26c-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26c-169">1.0</span><span class="sxs-lookup"><span data-stu-id="be26c-169">1.0</span></span>|
|[<span data-ttu-id="be26c-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be26c-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26c-171">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="be26c-171">Compose or read</span></span>|