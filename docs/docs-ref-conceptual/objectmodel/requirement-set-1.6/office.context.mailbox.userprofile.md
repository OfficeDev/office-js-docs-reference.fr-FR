
# <a name="userprofile"></a><span data-ttu-id="76090-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="76090-101">userProfile</span></span>

### <span data-ttu-id="76090-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="76090-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="76090-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76090-104">Requirements</span></span>

|<span data-ttu-id="76090-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76090-105">Requirement</span></span>| <span data-ttu-id="76090-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="76090-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="76090-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76090-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76090-108">1.0</span><span class="sxs-lookup"><span data-stu-id="76090-108">1.0</span></span>|
|[<span data-ttu-id="76090-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76090-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76090-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76090-110">ReadItem</span></span>|
|[<span data-ttu-id="76090-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76090-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76090-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="76090-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="76090-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="76090-113">Members and methods</span></span>

| <span data-ttu-id="76090-114">Membre</span><span class="sxs-lookup"><span data-stu-id="76090-114">Member</span></span> | <span data-ttu-id="76090-115">Type</span><span class="sxs-lookup"><span data-stu-id="76090-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="76090-116">accountType</span><span class="sxs-lookup"><span data-stu-id="76090-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="76090-117">Member</span><span class="sxs-lookup"><span data-stu-id="76090-117">Member</span></span> |
| [<span data-ttu-id="76090-118">displayName</span><span class="sxs-lookup"><span data-stu-id="76090-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="76090-119">Membre</span><span class="sxs-lookup"><span data-stu-id="76090-119">Member</span></span> |
| [<span data-ttu-id="76090-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="76090-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="76090-121">Membre</span><span class="sxs-lookup"><span data-stu-id="76090-121">Member</span></span> |
| [<span data-ttu-id="76090-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="76090-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="76090-123">Membre</span><span class="sxs-lookup"><span data-stu-id="76090-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="76090-124">Members</span><span class="sxs-lookup"><span data-stu-id="76090-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="76090-125">accountType : chaîne</span><span class="sxs-lookup"><span data-stu-id="76090-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="76090-126">Ce membre est actuellement uniquement pris en charge dans Outlook 2016 ou version ultérieure pour Mac (build 16.9.1212 ou version ultérieure).</span><span class="sxs-lookup"><span data-stu-id="76090-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="76090-127">Obtient le type de compte de l’utilisateur associé à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="76090-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="76090-128">Les valeurs possibles sont répertoriés dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="76090-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="76090-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="76090-129">Value</span></span> | <span data-ttu-id="76090-130">Description</span><span class="sxs-lookup"><span data-stu-id="76090-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="76090-131">La boîte aux lettres est sur un serveur d’Exchange sur site.</span><span class="sxs-lookup"><span data-stu-id="76090-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="76090-132">La boîte aux lettres est associé à un compte Gmail.</span><span class="sxs-lookup"><span data-stu-id="76090-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="76090-133">La boîte aux lettres est associé avec un Office 365 fonctionne ou école compte.</span><span class="sxs-lookup"><span data-stu-id="76090-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="76090-134">La boîte aux lettres est associé à un compte Outlook.com personnel.</span><span class="sxs-lookup"><span data-stu-id="76090-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="76090-135">Type :</span><span class="sxs-lookup"><span data-stu-id="76090-135">Type:</span></span>

*   <span data-ttu-id="76090-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="76090-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="76090-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76090-137">Requirements</span></span>

|<span data-ttu-id="76090-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76090-138">Requirement</span></span>| <span data-ttu-id="76090-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="76090-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="76090-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76090-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76090-141">1.6</span><span class="sxs-lookup"><span data-stu-id="76090-141">1.6</span></span> |
|[<span data-ttu-id="76090-142">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76090-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76090-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76090-143">ReadItem</span></span>|
|[<span data-ttu-id="76090-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76090-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76090-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="76090-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="76090-146">Exemple</span><span class="sxs-lookup"><span data-stu-id="76090-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="76090-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="76090-147">displayName :String</span></span>

<span data-ttu-id="76090-148">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="76090-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="76090-149">Type :</span><span class="sxs-lookup"><span data-stu-id="76090-149">Type:</span></span>

*   <span data-ttu-id="76090-150">Chaîne</span><span class="sxs-lookup"><span data-stu-id="76090-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="76090-151">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76090-151">Requirements</span></span>

|<span data-ttu-id="76090-152">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76090-152">Requirement</span></span>| <span data-ttu-id="76090-153">Valeur</span><span class="sxs-lookup"><span data-stu-id="76090-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="76090-154">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76090-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76090-155">1.0</span><span class="sxs-lookup"><span data-stu-id="76090-155">1.0</span></span>|
|[<span data-ttu-id="76090-156">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76090-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76090-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76090-157">ReadItem</span></span>|
|[<span data-ttu-id="76090-158">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76090-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76090-159">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="76090-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="76090-160">Exemple</span><span class="sxs-lookup"><span data-stu-id="76090-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="76090-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="76090-161">emailAddress :String</span></span>

<span data-ttu-id="76090-162">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="76090-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="76090-163">Type :</span><span class="sxs-lookup"><span data-stu-id="76090-163">Type:</span></span>

*   <span data-ttu-id="76090-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="76090-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="76090-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76090-165">Requirements</span></span>

|<span data-ttu-id="76090-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76090-166">Requirement</span></span>| <span data-ttu-id="76090-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="76090-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="76090-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76090-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76090-169">1.0</span><span class="sxs-lookup"><span data-stu-id="76090-169">1.0</span></span>|
|[<span data-ttu-id="76090-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76090-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76090-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76090-171">ReadItem</span></span>|
|[<span data-ttu-id="76090-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76090-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76090-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="76090-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="76090-174">Exemple</span><span class="sxs-lookup"><span data-stu-id="76090-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="76090-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="76090-175">timeZone :String</span></span>

<span data-ttu-id="76090-176">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="76090-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="76090-177">Type :</span><span class="sxs-lookup"><span data-stu-id="76090-177">Type:</span></span>

*   <span data-ttu-id="76090-178">Chaîne</span><span class="sxs-lookup"><span data-stu-id="76090-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="76090-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="76090-179">Requirements</span></span>

|<span data-ttu-id="76090-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="76090-180">Requirement</span></span>| <span data-ttu-id="76090-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="76090-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="76090-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="76090-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76090-183">1.0</span><span class="sxs-lookup"><span data-stu-id="76090-183">1.0</span></span>|
|[<span data-ttu-id="76090-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="76090-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76090-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76090-185">ReadItem</span></span>|
|[<span data-ttu-id="76090-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="76090-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76090-187">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="76090-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="76090-188">Exemples</span><span class="sxs-lookup"><span data-stu-id="76090-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```