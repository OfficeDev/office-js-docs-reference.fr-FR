 

# <a name="office"></a><span data-ttu-id="9e360-101">Bureau</span><span class="sxs-lookup"><span data-stu-id="9e360-101">Office</span></span>

<span data-ttu-id="9e360-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="9e360-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e360-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9e360-104">Requirements</span></span>

|<span data-ttu-id="9e360-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9e360-105">Requirement</span></span>| <span data-ttu-id="9e360-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="9e360-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e360-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9e360-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e360-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9e360-108">1.0</span></span>|
|[<span data-ttu-id="9e360-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9e360-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e360-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9e360-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="9e360-111">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="9e360-111">Namespaces</span></span>

<span data-ttu-id="9e360-112">[contexte](Office.context.md): fournit des interfaces partagées d’espace de noms de contexte au bureau, l’API de compléments pour une utilisation dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="9e360-112">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="9e360-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="9e360-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="9e360-114">Membres</span><span class="sxs-lookup"><span data-stu-id="9e360-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="9e360-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="9e360-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="9e360-116">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="9e360-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9e360-117">Type :</span><span class="sxs-lookup"><span data-stu-id="9e360-117">Type:</span></span>

*   <span data-ttu-id="9e360-118">String</span><span class="sxs-lookup"><span data-stu-id="9e360-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9e360-119">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="9e360-119">Properties:</span></span>

|<span data-ttu-id="9e360-120">Nom</span><span class="sxs-lookup"><span data-stu-id="9e360-120">Name</span></span>| <span data-ttu-id="9e360-121">Type</span><span class="sxs-lookup"><span data-stu-id="9e360-121">Type</span></span>| <span data-ttu-id="9e360-122">Description</span><span class="sxs-lookup"><span data-stu-id="9e360-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9e360-123">String</span><span class="sxs-lookup"><span data-stu-id="9e360-123">String</span></span>|<span data-ttu-id="9e360-124">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="9e360-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9e360-125">String</span><span class="sxs-lookup"><span data-stu-id="9e360-125">String</span></span>|<span data-ttu-id="9e360-126">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="9e360-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e360-127">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9e360-127">Requirements</span></span>

|<span data-ttu-id="9e360-128">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9e360-128">Requirement</span></span>| <span data-ttu-id="9e360-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="9e360-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e360-130">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9e360-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e360-131">1.0</span><span class="sxs-lookup"><span data-stu-id="9e360-131">1.0</span></span>|
|[<span data-ttu-id="9e360-132">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9e360-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e360-133">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9e360-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="9e360-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="9e360-134">CoercionType :String</span></span>

<span data-ttu-id="9e360-135">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="9e360-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9e360-136">Type :</span><span class="sxs-lookup"><span data-stu-id="9e360-136">Type:</span></span>

*   <span data-ttu-id="9e360-137">String</span><span class="sxs-lookup"><span data-stu-id="9e360-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9e360-138">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="9e360-138">Properties:</span></span>

|<span data-ttu-id="9e360-139">Nom</span><span class="sxs-lookup"><span data-stu-id="9e360-139">Name</span></span>| <span data-ttu-id="9e360-140">Type</span><span class="sxs-lookup"><span data-stu-id="9e360-140">Type</span></span>| <span data-ttu-id="9e360-141">Description</span><span class="sxs-lookup"><span data-stu-id="9e360-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9e360-142">String</span><span class="sxs-lookup"><span data-stu-id="9e360-142">String</span></span>|<span data-ttu-id="9e360-143">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="9e360-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9e360-144">String</span><span class="sxs-lookup"><span data-stu-id="9e360-144">String</span></span>|<span data-ttu-id="9e360-145">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="9e360-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e360-146">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9e360-146">Requirements</span></span>

|<span data-ttu-id="9e360-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9e360-147">Requirement</span></span>| <span data-ttu-id="9e360-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="9e360-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e360-149">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9e360-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e360-150">1.0</span><span class="sxs-lookup"><span data-stu-id="9e360-150">1.0</span></span>|
|[<span data-ttu-id="9e360-151">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9e360-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e360-152">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9e360-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="9e360-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="9e360-153">SourceProperty :String</span></span>

<span data-ttu-id="9e360-154">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="9e360-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9e360-155">Type :</span><span class="sxs-lookup"><span data-stu-id="9e360-155">Type:</span></span>

*   <span data-ttu-id="9e360-156">String</span><span class="sxs-lookup"><span data-stu-id="9e360-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9e360-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="9e360-157">Properties:</span></span>

|<span data-ttu-id="9e360-158">Nom</span><span class="sxs-lookup"><span data-stu-id="9e360-158">Name</span></span>| <span data-ttu-id="9e360-159">Type</span><span class="sxs-lookup"><span data-stu-id="9e360-159">Type</span></span>| <span data-ttu-id="9e360-160">Description</span><span class="sxs-lookup"><span data-stu-id="9e360-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9e360-161">String</span><span class="sxs-lookup"><span data-stu-id="9e360-161">String</span></span>|<span data-ttu-id="9e360-162">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="9e360-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9e360-163">String</span><span class="sxs-lookup"><span data-stu-id="9e360-163">String</span></span>|<span data-ttu-id="9e360-164">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="9e360-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e360-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9e360-165">Requirements</span></span>

|<span data-ttu-id="9e360-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9e360-166">Requirement</span></span>| <span data-ttu-id="9e360-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="9e360-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e360-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9e360-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e360-169">1.0</span><span class="sxs-lookup"><span data-stu-id="9e360-169">1.0</span></span>|
|[<span data-ttu-id="9e360-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9e360-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9e360-171">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="9e360-171">Compose or read</span></span>|