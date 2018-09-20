# <a name="userprofile"></a><span data-ttu-id="978a2-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="978a2-101">userProfile</span></span>

### <span data-ttu-id="978a2-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="978a2-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="978a2-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="978a2-104">Requirements</span></span>

|<span data-ttu-id="978a2-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="978a2-105">Requirement</span></span>| <span data-ttu-id="978a2-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="978a2-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="978a2-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="978a2-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="978a2-108">1.0</span><span class="sxs-lookup"><span data-stu-id="978a2-108">1.0</span></span>|
|[<span data-ttu-id="978a2-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="978a2-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="978a2-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="978a2-110">ReadItem</span></span>|
|[<span data-ttu-id="978a2-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="978a2-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="978a2-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="978a2-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="978a2-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="978a2-113">Members and methods</span></span>

| <span data-ttu-id="978a2-114">Membre</span><span class="sxs-lookup"><span data-stu-id="978a2-114">Member</span></span> | <span data-ttu-id="978a2-115">Type</span><span class="sxs-lookup"><span data-stu-id="978a2-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="978a2-116">displayName</span><span class="sxs-lookup"><span data-stu-id="978a2-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="978a2-117">Membre</span><span class="sxs-lookup"><span data-stu-id="978a2-117">Member</span></span> |
| [<span data-ttu-id="978a2-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="978a2-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="978a2-119">Membre</span><span class="sxs-lookup"><span data-stu-id="978a2-119">Member</span></span> |
| [<span data-ttu-id="978a2-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="978a2-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="978a2-121">Membre</span><span class="sxs-lookup"><span data-stu-id="978a2-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="978a2-122">Membres</span><span class="sxs-lookup"><span data-stu-id="978a2-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="978a2-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="978a2-123">displayName :String</span></span>

<span data-ttu-id="978a2-124">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="978a2-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="978a2-125">Type :</span><span class="sxs-lookup"><span data-stu-id="978a2-125">Type:</span></span>

*   <span data-ttu-id="978a2-126">Chaîne</span><span class="sxs-lookup"><span data-stu-id="978a2-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="978a2-127">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="978a2-127">Requirements</span></span>

|<span data-ttu-id="978a2-128">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="978a2-128">Requirement</span></span>| <span data-ttu-id="978a2-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="978a2-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="978a2-130">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="978a2-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="978a2-131">1.0</span><span class="sxs-lookup"><span data-stu-id="978a2-131">1.0</span></span>|
|[<span data-ttu-id="978a2-132">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="978a2-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="978a2-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="978a2-133">ReadItem</span></span>|
|[<span data-ttu-id="978a2-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="978a2-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="978a2-135">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="978a2-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="978a2-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="978a2-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="978a2-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="978a2-137">emailAddress :String</span></span>

<span data-ttu-id="978a2-138">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="978a2-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="978a2-139">Type :</span><span class="sxs-lookup"><span data-stu-id="978a2-139">Type:</span></span>

*   <span data-ttu-id="978a2-140">Chaîne</span><span class="sxs-lookup"><span data-stu-id="978a2-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="978a2-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="978a2-141">Requirements</span></span>

|<span data-ttu-id="978a2-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="978a2-142">Requirement</span></span>| <span data-ttu-id="978a2-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="978a2-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="978a2-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="978a2-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="978a2-145">1.0</span><span class="sxs-lookup"><span data-stu-id="978a2-145">1.0</span></span>|
|[<span data-ttu-id="978a2-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="978a2-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="978a2-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="978a2-147">ReadItem</span></span>|
|[<span data-ttu-id="978a2-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="978a2-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="978a2-149">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="978a2-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="978a2-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="978a2-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="978a2-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="978a2-151">timeZone :String</span></span>

<span data-ttu-id="978a2-152">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="978a2-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="978a2-153">Type :</span><span class="sxs-lookup"><span data-stu-id="978a2-153">Type:</span></span>

*   <span data-ttu-id="978a2-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="978a2-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="978a2-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="978a2-155">Requirements</span></span>

|<span data-ttu-id="978a2-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="978a2-156">Requirement</span></span>| <span data-ttu-id="978a2-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="978a2-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="978a2-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="978a2-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="978a2-159">1.0</span><span class="sxs-lookup"><span data-stu-id="978a2-159">1.0</span></span>|
|[<span data-ttu-id="978a2-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="978a2-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="978a2-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="978a2-161">ReadItem</span></span>|
|[<span data-ttu-id="978a2-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="978a2-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="978a2-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="978a2-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="978a2-164">Exemples</span><span class="sxs-lookup"><span data-stu-id="978a2-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```