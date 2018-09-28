
# <a name="userprofile"></a><span data-ttu-id="b9dd9-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="b9dd9-101">userProfile</span></span>

### <span data-ttu-id="b9dd9-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="b9dd9-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9dd9-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9dd9-104">Requirements</span></span>

|<span data-ttu-id="b9dd9-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9dd9-105">Requirement</span></span>| <span data-ttu-id="b9dd9-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9dd9-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9dd9-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9dd9-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9dd9-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b9dd9-108">1.0</span></span>|
|[<span data-ttu-id="b9dd9-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9dd9-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9dd9-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9dd9-110">ReadItem</span></span>|
|[<span data-ttu-id="b9dd9-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9dd9-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9dd9-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9dd9-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="b9dd9-113">Membres</span><span class="sxs-lookup"><span data-stu-id="b9dd9-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="b9dd9-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="b9dd9-114">displayName :String</span></span>

<span data-ttu-id="b9dd9-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b9dd9-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="b9dd9-116">Type :</span><span class="sxs-lookup"><span data-stu-id="b9dd9-116">Type:</span></span>

*   <span data-ttu-id="b9dd9-117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b9dd9-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9dd9-118">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9dd9-118">Requirements</span></span>

|<span data-ttu-id="b9dd9-119">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9dd9-119">Requirement</span></span>| <span data-ttu-id="b9dd9-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9dd9-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9dd9-121">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9dd9-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9dd9-122">1.0</span><span class="sxs-lookup"><span data-stu-id="b9dd9-122">1.0</span></span>|
|[<span data-ttu-id="b9dd9-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9dd9-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9dd9-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9dd9-124">ReadItem</span></span>|
|[<span data-ttu-id="b9dd9-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9dd9-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9dd9-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9dd9-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9dd9-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9dd9-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="b9dd9-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="b9dd9-128">emailAddress :String</span></span>

<span data-ttu-id="b9dd9-129">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b9dd9-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="b9dd9-130">Type :</span><span class="sxs-lookup"><span data-stu-id="b9dd9-130">Type:</span></span>

*   <span data-ttu-id="b9dd9-131">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b9dd9-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9dd9-132">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9dd9-132">Requirements</span></span>

|<span data-ttu-id="b9dd9-133">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9dd9-133">Requirement</span></span>| <span data-ttu-id="b9dd9-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9dd9-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9dd9-135">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9dd9-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9dd9-136">1.0</span><span class="sxs-lookup"><span data-stu-id="b9dd9-136">1.0</span></span>|
|[<span data-ttu-id="b9dd9-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9dd9-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9dd9-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9dd9-138">ReadItem</span></span>|
|[<span data-ttu-id="b9dd9-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9dd9-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9dd9-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9dd9-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9dd9-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="b9dd9-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="b9dd9-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="b9dd9-142">timeZone :String</span></span>

<span data-ttu-id="b9dd9-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b9dd9-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="b9dd9-144">Type :</span><span class="sxs-lookup"><span data-stu-id="b9dd9-144">Type:</span></span>

*   <span data-ttu-id="b9dd9-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b9dd9-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9dd9-146">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b9dd9-146">Requirements</span></span>

|<span data-ttu-id="b9dd9-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b9dd9-147">Requirement</span></span>| <span data-ttu-id="b9dd9-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="b9dd9-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9dd9-149">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b9dd9-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9dd9-150">1.0</span><span class="sxs-lookup"><span data-stu-id="b9dd9-150">1.0</span></span>|
|[<span data-ttu-id="b9dd9-151">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b9dd9-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9dd9-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9dd9-152">ReadItem</span></span>|
|[<span data-ttu-id="b9dd9-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b9dd9-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9dd9-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="b9dd9-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9dd9-155">Exemples</span><span class="sxs-lookup"><span data-stu-id="b9dd9-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```