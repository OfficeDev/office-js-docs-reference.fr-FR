
# <a name="userprofile"></a><span data-ttu-id="bc9a8-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="bc9a8-101">userProfile</span></span>

### <span data-ttu-id="bc9a8-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="bc9a8-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="bc9a8-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bc9a8-104">Requirements</span></span>

|<span data-ttu-id="bc9a8-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc9a8-105">Requirement</span></span>| <span data-ttu-id="bc9a8-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc9a8-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc9a8-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bc9a8-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc9a8-108">1.0</span><span class="sxs-lookup"><span data-stu-id="bc9a8-108">1.0</span></span>|
|[<span data-ttu-id="bc9a8-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bc9a8-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc9a8-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc9a8-110">ReadItem</span></span>|
|[<span data-ttu-id="bc9a8-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bc9a8-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bc9a8-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="bc9a8-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bc9a8-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="bc9a8-113">Members and methods</span></span>

| <span data-ttu-id="bc9a8-114">Membre</span><span class="sxs-lookup"><span data-stu-id="bc9a8-114">Member</span></span> | <span data-ttu-id="bc9a8-115">Type</span><span class="sxs-lookup"><span data-stu-id="bc9a8-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bc9a8-116">accountType</span><span class="sxs-lookup"><span data-stu-id="bc9a8-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="bc9a8-117">Member</span><span class="sxs-lookup"><span data-stu-id="bc9a8-117">Member</span></span> |
| [<span data-ttu-id="bc9a8-118">displayName</span><span class="sxs-lookup"><span data-stu-id="bc9a8-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="bc9a8-119">Membre</span><span class="sxs-lookup"><span data-stu-id="bc9a8-119">Member</span></span> |
| [<span data-ttu-id="bc9a8-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="bc9a8-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="bc9a8-121">Membre</span><span class="sxs-lookup"><span data-stu-id="bc9a8-121">Member</span></span> |
| [<span data-ttu-id="bc9a8-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="bc9a8-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="bc9a8-123">Membre</span><span class="sxs-lookup"><span data-stu-id="bc9a8-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="bc9a8-124">Members</span><span class="sxs-lookup"><span data-stu-id="bc9a8-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="bc9a8-125">accountType : chaîne</span><span class="sxs-lookup"><span data-stu-id="bc9a8-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="bc9a8-126">Ce membre est uniquement pris en charge dans 2016 Outlook pour Mac, créer 16.9.1212 et versions ultérieures.</span><span class="sxs-lookup"><span data-stu-id="bc9a8-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="bc9a8-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="bc9a8-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="bc9a8-128">Les valeurs possibles sont répertoriés dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="bc9a8-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="bc9a8-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc9a8-129">Value</span></span> | <span data-ttu-id="bc9a8-130">Description</span><span class="sxs-lookup"><span data-stu-id="bc9a8-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="bc9a8-131">La boîte aux lettres est sur un serveur d’Exchange sur site.</span><span class="sxs-lookup"><span data-stu-id="bc9a8-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="bc9a8-132">La boîte aux lettres est associé à un compte Gmail.</span><span class="sxs-lookup"><span data-stu-id="bc9a8-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="bc9a8-133">La boîte aux lettres est associé avec un Office 365 fonctionne ou école compte.</span><span class="sxs-lookup"><span data-stu-id="bc9a8-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="bc9a8-134">La boîte aux lettres est associé à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="bc9a8-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="bc9a8-135">Type :</span><span class="sxs-lookup"><span data-stu-id="bc9a8-135">Type:</span></span>

*   <span data-ttu-id="bc9a8-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bc9a8-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bc9a8-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bc9a8-137">Requirements</span></span>

|<span data-ttu-id="bc9a8-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc9a8-138">Requirement</span></span>| <span data-ttu-id="bc9a8-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc9a8-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc9a8-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bc9a8-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc9a8-141">1.6</span><span class="sxs-lookup"><span data-stu-id="bc9a8-141">1.6</span></span> |
|[<span data-ttu-id="bc9a8-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bc9a8-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc9a8-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc9a8-143">ReadItem</span></span>|
|[<span data-ttu-id="bc9a8-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bc9a8-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bc9a8-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="bc9a8-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bc9a8-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="bc9a8-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="bc9a8-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="bc9a8-147">displayName :String</span></span>

<span data-ttu-id="bc9a8-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="bc9a8-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="bc9a8-149">Type :</span><span class="sxs-lookup"><span data-stu-id="bc9a8-149">Type:</span></span>

*   <span data-ttu-id="bc9a8-150">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bc9a8-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bc9a8-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bc9a8-151">Requirements</span></span>

|<span data-ttu-id="bc9a8-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc9a8-152">Requirement</span></span>| <span data-ttu-id="bc9a8-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc9a8-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc9a8-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bc9a8-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc9a8-155">1.0</span><span class="sxs-lookup"><span data-stu-id="bc9a8-155">1.0</span></span>|
|[<span data-ttu-id="bc9a8-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bc9a8-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc9a8-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc9a8-157">ReadItem</span></span>|
|[<span data-ttu-id="bc9a8-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bc9a8-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bc9a8-159">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="bc9a8-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bc9a8-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="bc9a8-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="bc9a8-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="bc9a8-161">emailAddress :String</span></span>

<span data-ttu-id="bc9a8-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="bc9a8-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="bc9a8-163">Type :</span><span class="sxs-lookup"><span data-stu-id="bc9a8-163">Type:</span></span>

*   <span data-ttu-id="bc9a8-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bc9a8-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bc9a8-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bc9a8-165">Requirements</span></span>

|<span data-ttu-id="bc9a8-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc9a8-166">Requirement</span></span>| <span data-ttu-id="bc9a8-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc9a8-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc9a8-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bc9a8-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc9a8-169">1.0</span><span class="sxs-lookup"><span data-stu-id="bc9a8-169">1.0</span></span>|
|[<span data-ttu-id="bc9a8-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bc9a8-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc9a8-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc9a8-171">ReadItem</span></span>|
|[<span data-ttu-id="bc9a8-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bc9a8-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bc9a8-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="bc9a8-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bc9a8-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="bc9a8-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="bc9a8-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="bc9a8-175">timeZone :String</span></span>

<span data-ttu-id="bc9a8-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="bc9a8-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="bc9a8-177">Type :</span><span class="sxs-lookup"><span data-stu-id="bc9a8-177">Type:</span></span>

*   <span data-ttu-id="bc9a8-178">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bc9a8-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bc9a8-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bc9a8-179">Requirements</span></span>

|<span data-ttu-id="bc9a8-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc9a8-180">Requirement</span></span>| <span data-ttu-id="bc9a8-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc9a8-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc9a8-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bc9a8-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bc9a8-183">1.0</span><span class="sxs-lookup"><span data-stu-id="bc9a8-183">1.0</span></span>|
|[<span data-ttu-id="bc9a8-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="bc9a8-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bc9a8-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bc9a8-185">ReadItem</span></span>|
|[<span data-ttu-id="bc9a8-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bc9a8-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="bc9a8-187">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="bc9a8-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="bc9a8-188">Exemples</span><span class="sxs-lookup"><span data-stu-id="bc9a8-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```