
# <a name="item"></a><span data-ttu-id="f39c6-101">élément</span><span class="sxs-lookup"><span data-stu-id="f39c6-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="f39c6-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f39c6-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="f39c6-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="f39c6-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-105">Requirements</span></span>

|<span data-ttu-id="f39c6-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-106">Requirement</span></span>| <span data-ttu-id="f39c6-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-109">1.0</span></span>|
|[<span data-ttu-id="f39c6-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-111">Restricted</span><span class="sxs-lookup"><span data-stu-id="f39c6-111">Restricted</span></span>|
|[<span data-ttu-id="f39c6-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f39c6-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="f39c6-114">Members and methods</span></span>

| <span data-ttu-id="f39c6-115">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-115">Member</span></span> | <span data-ttu-id="f39c6-116">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f39c6-117">attachments</span><span class="sxs-lookup"><span data-stu-id="f39c6-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="f39c6-118">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-118">Member</span></span> |
| [<span data-ttu-id="f39c6-119">bcc</span><span class="sxs-lookup"><span data-stu-id="f39c6-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="f39c6-120">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-120">Member</span></span> |
| [<span data-ttu-id="f39c6-121">body</span><span class="sxs-lookup"><span data-stu-id="f39c6-121">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="f39c6-122">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-122">Member</span></span> |
| [<span data-ttu-id="f39c6-123">cc</span><span class="sxs-lookup"><span data-stu-id="f39c6-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="f39c6-124">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-124">Member</span></span> |
| [<span data-ttu-id="f39c6-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="f39c6-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="f39c6-126">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-126">Member</span></span> |
| [<span data-ttu-id="f39c6-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="f39c6-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="f39c6-128">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-128">Member</span></span> |
| [<span data-ttu-id="f39c6-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="f39c6-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="f39c6-130">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-130">Member</span></span> |
| [<span data-ttu-id="f39c6-131">end</span><span class="sxs-lookup"><span data-stu-id="f39c6-131">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="f39c6-132">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-132">Member</span></span> |
| [<span data-ttu-id="f39c6-133">from</span><span class="sxs-lookup"><span data-stu-id="f39c6-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="f39c6-134">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-134">Member</span></span> |
| [<span data-ttu-id="f39c6-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="f39c6-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="f39c6-136">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-136">Member</span></span> |
| [<span data-ttu-id="f39c6-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="f39c6-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="f39c6-138">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-138">Member</span></span> |
| [<span data-ttu-id="f39c6-139">itemId</span><span class="sxs-lookup"><span data-stu-id="f39c6-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="f39c6-140">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-140">Member</span></span> |
| [<span data-ttu-id="f39c6-141">itemType</span><span class="sxs-lookup"><span data-stu-id="f39c6-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="f39c6-142">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-142">Member</span></span> |
| [<span data-ttu-id="f39c6-143">location</span><span class="sxs-lookup"><span data-stu-id="f39c6-143">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="f39c6-144">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-144">Member</span></span> |
| [<span data-ttu-id="f39c6-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="f39c6-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="f39c6-146">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-146">Member</span></span> |
| [<span data-ttu-id="f39c6-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="f39c6-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="f39c6-148">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-148">Member</span></span> |
| [<span data-ttu-id="f39c6-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="f39c6-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="f39c6-150">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-150">Member</span></span> |
| [<span data-ttu-id="f39c6-151">organizer</span><span class="sxs-lookup"><span data-stu-id="f39c6-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="f39c6-152">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-152">Member</span></span> |
| [<span data-ttu-id="f39c6-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="f39c6-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="f39c6-154">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-154">Member</span></span> |
| [<span data-ttu-id="f39c6-155">sender</span><span class="sxs-lookup"><span data-stu-id="f39c6-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="f39c6-156">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-156">Member</span></span> |
| [<span data-ttu-id="f39c6-157">start</span><span class="sxs-lookup"><span data-stu-id="f39c6-157">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="f39c6-158">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-158">Member</span></span> |
| [<span data-ttu-id="f39c6-159">subject</span><span class="sxs-lookup"><span data-stu-id="f39c6-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="f39c6-160">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-160">Member</span></span> |
| [<span data-ttu-id="f39c6-161">to</span><span class="sxs-lookup"><span data-stu-id="f39c6-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="f39c6-162">Membre</span><span class="sxs-lookup"><span data-stu-id="f39c6-162">Member</span></span> |
| [<span data-ttu-id="f39c6-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f39c6-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="f39c6-164">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-164">Method</span></span> |
| [<span data-ttu-id="f39c6-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f39c6-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="f39c6-166">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-166">Method</span></span> |
| [<span data-ttu-id="f39c6-167">close</span><span class="sxs-lookup"><span data-stu-id="f39c6-167">close</span></span>](#close) | <span data-ttu-id="f39c6-168">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-168">Method</span></span> |
| [<span data-ttu-id="f39c6-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="f39c6-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="f39c6-170">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-170">Method</span></span> |
| [<span data-ttu-id="f39c6-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="f39c6-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="f39c6-172">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-172">Method</span></span> |
| [<span data-ttu-id="f39c6-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="f39c6-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="f39c6-174">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-174">Method</span></span> |
| [<span data-ttu-id="f39c6-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="f39c6-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="f39c6-176">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-176">Method</span></span> |
| [<span data-ttu-id="f39c6-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="f39c6-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="f39c6-178">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-178">Method</span></span> |
| [<span data-ttu-id="f39c6-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="f39c6-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="f39c6-180">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-180">Method</span></span> |
| [<span data-ttu-id="f39c6-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="f39c6-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="f39c6-182">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-182">Method</span></span> |
| [<span data-ttu-id="f39c6-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f39c6-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="f39c6-184">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-184">Method</span></span> |
| [<span data-ttu-id="f39c6-185">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f39c6-185">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="f39c6-186">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-186">Method</span></span> |
| [<span data-ttu-id="f39c6-187">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f39c6-187">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="f39c6-188">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-188">Method</span></span> |
| [<span data-ttu-id="f39c6-189">saveAsync</span><span class="sxs-lookup"><span data-stu-id="f39c6-189">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="f39c6-190">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-190">Method</span></span> |
| [<span data-ttu-id="f39c6-191">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f39c6-191">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="f39c6-192">Méthode</span><span class="sxs-lookup"><span data-stu-id="f39c6-192">Method</span></span> |

### <a name="example"></a><span data-ttu-id="f39c6-193">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-193">Example</span></span>

<span data-ttu-id="f39c6-194">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="f39c6-194">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="f39c6-195">Membres</span><span class="sxs-lookup"><span data-stu-id="f39c6-195">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="f39c6-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f39c6-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="f39c6-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-199">Certains types de fichiers bloqués par Outlook en raison de problèmes de sécurité potentiels et ne sont donc pas retournés.</span><span class="sxs-lookup"><span data-stu-id="f39c6-199">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f39c6-200">Pour plus d’informations, voir les [pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="f39c6-200">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-201">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-201">Type:</span></span>

*   <span data-ttu-id="f39c6-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f39c6-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-203">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-203">Requirements</span></span>

|<span data-ttu-id="f39c6-204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-204">Requirement</span></span>| <span data-ttu-id="f39c6-205">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-206">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-206">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-207">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-207">1.0</span></span>|
|[<span data-ttu-id="f39c6-208">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-209">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-210">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-211">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-211">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-212">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-212">Example</span></span>

<span data-ttu-id="f39c6-213">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="f39c6-213">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="f39c6-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f39c6-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="f39c6-215">Obtient un objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires des Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="f39c6-215">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f39c6-216">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="f39c6-216">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-217">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-217">Type:</span></span>

*   [<span data-ttu-id="f39c6-218">Destinataires</span><span class="sxs-lookup"><span data-stu-id="f39c6-218">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="f39c6-219">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-219">Requirements</span></span>

|<span data-ttu-id="f39c6-220">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-220">Requirement</span></span>| <span data-ttu-id="f39c6-221">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-222">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-222">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-223">1.1</span><span class="sxs-lookup"><span data-stu-id="f39c6-223">1.1</span></span>|
|[<span data-ttu-id="f39c6-224">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-225">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-226">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-227">Composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-228">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-228">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="f39c6-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="f39c6-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="f39c6-230">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f39c6-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-231">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-231">Type:</span></span>

*   [<span data-ttu-id="f39c6-232">Corps</span><span class="sxs-lookup"><span data-stu-id="f39c6-232">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="f39c6-233">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-233">Requirements</span></span>

|<span data-ttu-id="f39c6-234">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-234">Requirement</span></span>| <span data-ttu-id="f39c6-235">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-236">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-236">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-237">1.1</span><span class="sxs-lookup"><span data-stu-id="f39c6-237">1.1</span></span>|
|[<span data-ttu-id="f39c6-238">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-239">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-240">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-241">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-241">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="f39c6-242">cc : tableau. <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f39c6-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="f39c6-243">Permet d’accéder aux destinataires Cc (copie carbone) d’un message.</span><span class="sxs-lookup"><span data-stu-id="f39c6-243">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f39c6-244">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-244">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f39c6-245">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-245">Read mode</span></span>

<span data-ttu-id="f39c6-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f39c6-248">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-248">Compose mode</span></span>

<span data-ttu-id="f39c6-249">Le `cc` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="f39c6-249">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-250">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-250">Type:</span></span>

*   <span data-ttu-id="f39c6-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f39c6-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-252">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-252">Requirements</span></span>

|<span data-ttu-id="f39c6-253">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-253">Requirement</span></span>| <span data-ttu-id="f39c6-254">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-255">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-255">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-256">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-256">1.0</span></span>|
|[<span data-ttu-id="f39c6-257">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-257">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-258">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-259">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-259">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-260">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-260">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-261">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-261">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="f39c6-262">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="f39c6-262">(nullable) conversationId :String</span></span>

<span data-ttu-id="f39c6-263">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="f39c6-263">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f39c6-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f39c6-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-268">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-268">Type:</span></span>

*   <span data-ttu-id="f39c6-269">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f39c6-269">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-270">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-270">Requirements</span></span>

|<span data-ttu-id="f39c6-271">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-271">Requirement</span></span>| <span data-ttu-id="f39c6-272">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-273">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-274">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-274">1.0</span></span>|
|[<span data-ttu-id="f39c6-275">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-276">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-277">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-278">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-278">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="f39c6-279">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="f39c6-279">dateTimeCreated :Date</span></span>

<span data-ttu-id="f39c6-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-282">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-282">Type:</span></span>

*   <span data-ttu-id="f39c6-283">Date</span><span class="sxs-lookup"><span data-stu-id="f39c6-283">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-284">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-284">Requirements</span></span>

|<span data-ttu-id="f39c6-285">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-285">Requirement</span></span>| <span data-ttu-id="f39c6-286">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-287">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-287">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-288">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-288">1.0</span></span>|
|[<span data-ttu-id="f39c6-289">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-289">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-290">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-290">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-291">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-291">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-292">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-292">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-293">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-293">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="f39c6-294">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="f39c6-294">dateTimeModified :Date</span></span>

<span data-ttu-id="f39c6-p110">Obtient la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-297">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f39c6-297">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-298">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-298">Type:</span></span>

*   <span data-ttu-id="f39c6-299">Date</span><span class="sxs-lookup"><span data-stu-id="f39c6-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-300">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-300">Requirements</span></span>

|<span data-ttu-id="f39c6-301">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-301">Requirement</span></span>| <span data-ttu-id="f39c6-302">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-303">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-304">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-304">1.0</span></span>|
|[<span data-ttu-id="f39c6-305">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-306">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-307">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-308">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-309">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-309">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="f39c6-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="f39c6-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="f39c6-311">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f39c6-311">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f39c6-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f39c6-314">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-314">Read mode</span></span>

<span data-ttu-id="f39c6-315">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-315">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f39c6-316">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-316">Compose mode</span></span>

<span data-ttu-id="f39c6-317">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-317">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f39c6-318">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="f39c6-318">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-319">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-319">Type:</span></span>

*   <span data-ttu-id="f39c6-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="f39c6-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-321">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-321">Requirements</span></span>

|<span data-ttu-id="f39c6-322">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-322">Requirement</span></span>| <span data-ttu-id="f39c6-323">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-324">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-324">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-325">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-325">1.0</span></span>|
|[<span data-ttu-id="f39c6-326">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-327">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-328">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-329">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-329">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-330">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-330">Example</span></span>

<span data-ttu-id="f39c6-331">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-331">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="f39c6-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f39c6-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="f39c6-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="f39c6-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-337">Le `recipientType` propriété de la `EmailAddressDetails` objet dans le `from` propriété est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-337">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-338">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-338">Type:</span></span>

*   [<span data-ttu-id="f39c6-339">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f39c6-339">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f39c6-340">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-340">Requirements</span></span>

|<span data-ttu-id="f39c6-341">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-341">Requirement</span></span>| <span data-ttu-id="f39c6-342">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-342">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-343">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-343">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-344">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-344">1.0</span></span>|
|[<span data-ttu-id="f39c6-345">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-345">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-346">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-346">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-347">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-347">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-348">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-348">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="f39c6-349">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="f39c6-349">internetMessageId :String</span></span>

<span data-ttu-id="f39c6-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-352">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-352">Type:</span></span>

*   <span data-ttu-id="f39c6-353">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f39c6-353">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-354">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-354">Requirements</span></span>

|<span data-ttu-id="f39c6-355">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-355">Requirement</span></span>| <span data-ttu-id="f39c6-356">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-357">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-357">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-358">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-358">1.0</span></span>|
|[<span data-ttu-id="f39c6-359">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-360">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-361">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-362">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-362">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-363">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-363">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="f39c6-364">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="f39c6-364">itemClass :String</span></span>

<span data-ttu-id="f39c6-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f39c6-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="f39c6-369">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-369">Type</span></span> | <span data-ttu-id="f39c6-370">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-370">Description</span></span> | <span data-ttu-id="f39c6-371">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="f39c6-371">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="f39c6-372">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="f39c6-372">Appointment items</span></span> | <span data-ttu-id="f39c6-373">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-373">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="f39c6-374">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="f39c6-374">Message items</span></span> | <span data-ttu-id="f39c6-375">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="f39c6-375">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="f39c6-376">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-376">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-377">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-377">Type:</span></span>

*   <span data-ttu-id="f39c6-378">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f39c6-378">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-379">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-379">Requirements</span></span>

|<span data-ttu-id="f39c6-380">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-380">Requirement</span></span>| <span data-ttu-id="f39c6-381">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-381">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-382">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-382">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-383">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-383">1.0</span></span>|
|[<span data-ttu-id="f39c6-384">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-385">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-386">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-387">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-387">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-388">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-388">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f39c6-389">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="f39c6-389">(nullable) itemId :String</span></span>

<span data-ttu-id="f39c6-p117">Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-392">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="f39c6-392">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f39c6-393">Le `itemId` propriété n’est pas identique à l’identificateur d’entrée Outlook ou l’ID utilisé par l’API REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="f39c6-393">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f39c6-394">Avant d’effectuer des appels d’API REST à l’aide de cette valeur, elle doit être convertie à l’aide de [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="f39c6-394">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f39c6-395">Pour plus d’informations, voir [utiliser les API REST d’Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="f39c6-395">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="f39c6-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-398">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-398">Type:</span></span>

*   <span data-ttu-id="f39c6-399">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f39c6-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-400">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-400">Requirements</span></span>

|<span data-ttu-id="f39c6-401">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-401">Requirement</span></span>| <span data-ttu-id="f39c6-402">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-403">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-403">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-404">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-404">1.0</span></span>|
|[<span data-ttu-id="f39c6-405">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-406">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-407">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-408">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-409">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-409">Example</span></span>

<span data-ttu-id="f39c6-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="f39c6-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="f39c6-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="f39c6-413">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="f39c6-413">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f39c6-414">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f39c6-414">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-415">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-415">Type:</span></span>

*   [<span data-ttu-id="f39c6-416">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="f39c6-416">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="f39c6-417">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-417">Requirements</span></span>

|<span data-ttu-id="f39c6-418">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-418">Requirement</span></span>| <span data-ttu-id="f39c6-419">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-420">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-420">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-421">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-421">1.0</span></span>|
|[<span data-ttu-id="f39c6-422">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-423">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-424">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-425">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-425">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-426">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-426">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="f39c6-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="f39c6-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="f39c6-428">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f39c6-428">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f39c6-429">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-429">Read mode</span></span>

<span data-ttu-id="f39c6-430">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f39c6-430">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f39c6-431">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-431">Compose mode</span></span>

<span data-ttu-id="f39c6-432">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f39c6-432">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-433">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-433">Type:</span></span>

*   <span data-ttu-id="f39c6-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="f39c6-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-435">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-435">Requirements</span></span>

|<span data-ttu-id="f39c6-436">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-436">Requirement</span></span>| <span data-ttu-id="f39c6-437">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-438">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-438">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-439">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-439">1.0</span></span>|
|[<span data-ttu-id="f39c6-440">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-441">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-442">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-443">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-444">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-444">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f39c6-445">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="f39c6-445">normalizedSubject :String</span></span>

<span data-ttu-id="f39c6-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f39c6-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject).</span><span class="sxs-lookup"><span data-stu-id="f39c6-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-450">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-450">Type:</span></span>

*   <span data-ttu-id="f39c6-451">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f39c6-451">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-452">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-452">Requirements</span></span>

|<span data-ttu-id="f39c6-453">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-453">Requirement</span></span>| <span data-ttu-id="f39c6-454">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-455">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-455">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-456">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-456">1.0</span></span>|
|[<span data-ttu-id="f39c6-457">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-458">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-459">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-460">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-460">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-461">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-461">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="f39c6-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="f39c6-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="f39c6-463">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="f39c6-463">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-464">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-464">Type:</span></span>

*   [<span data-ttu-id="f39c6-465">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="f39c6-465">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="f39c6-466">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-466">Requirements</span></span>

|<span data-ttu-id="f39c6-467">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-467">Requirement</span></span>| <span data-ttu-id="f39c6-468">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-469">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-469">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-470">1.3</span><span class="sxs-lookup"><span data-stu-id="f39c6-470">1.3</span></span>|
|[<span data-ttu-id="f39c6-471">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-472">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-473">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-474">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-474">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="f39c6-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f39c6-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="f39c6-476">Fournit l’accès à un événement les participants facultatifs.</span><span class="sxs-lookup"><span data-stu-id="f39c6-476">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f39c6-477">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-477">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f39c6-478">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-478">Read mode</span></span>

<span data-ttu-id="f39c6-479">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="f39c6-479">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f39c6-480">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-480">Compose mode</span></span>

<span data-ttu-id="f39c6-481">Le `optionalAttendees` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs à une réunion.</span><span class="sxs-lookup"><span data-stu-id="f39c6-481">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-482">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-482">Type:</span></span>

*   <span data-ttu-id="f39c6-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f39c6-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-484">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-484">Requirements</span></span>

|<span data-ttu-id="f39c6-485">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-485">Requirement</span></span>| <span data-ttu-id="f39c6-486">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-487">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-487">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-488">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-488">1.0</span></span>|
|[<span data-ttu-id="f39c6-489">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-490">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-491">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-492">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-492">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-493">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-493">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="f39c6-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f39c6-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="f39c6-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-497">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-497">Type:</span></span>

*   [<span data-ttu-id="f39c6-498">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f39c6-498">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f39c6-499">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-499">Requirements</span></span>

|<span data-ttu-id="f39c6-500">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-500">Requirement</span></span>| <span data-ttu-id="f39c6-501">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-502">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-502">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-503">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-503">1.0</span></span>|
|[<span data-ttu-id="f39c6-504">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-505">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-506">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-507">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-508">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-508">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="f39c6-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f39c6-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="f39c6-510">Fournit l’accès à des participants obligatoires d’un événement.</span><span class="sxs-lookup"><span data-stu-id="f39c6-510">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f39c6-511">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f39c6-512">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-512">Read mode</span></span>

<span data-ttu-id="f39c6-513">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="f39c6-513">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f39c6-514">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-514">Compose mode</span></span>

<span data-ttu-id="f39c6-515">Le `requiredAttendees` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les participants obligatoires à une réunion.</span><span class="sxs-lookup"><span data-stu-id="f39c6-515">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-516">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-516">Type:</span></span>

*   <span data-ttu-id="f39c6-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f39c6-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-518">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-518">Requirements</span></span>

|<span data-ttu-id="f39c6-519">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-519">Requirement</span></span>| <span data-ttu-id="f39c6-520">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-521">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-522">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-522">1.0</span></span>|
|[<span data-ttu-id="f39c6-523">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-524">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-525">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-526">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-527">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-527">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="f39c6-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f39c6-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="f39c6-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f39c6-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-533">Le `recipientType` propriété de la `EmailAddressDetails` objet dans le `sender` propriété est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-533">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-534">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-534">Type:</span></span>

*   [<span data-ttu-id="f39c6-535">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f39c6-535">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f39c6-536">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-536">Requirements</span></span>

|<span data-ttu-id="f39c6-537">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-537">Requirement</span></span>| <span data-ttu-id="f39c6-538">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-539">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-539">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-540">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-540">1.0</span></span>|
|[<span data-ttu-id="f39c6-541">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-542">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-543">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-544">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-545">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-545">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="f39c6-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="f39c6-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="f39c6-547">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f39c6-547">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f39c6-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f39c6-550">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-550">Read mode</span></span>

<span data-ttu-id="f39c6-551">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-551">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f39c6-552">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-552">Compose mode</span></span>

<span data-ttu-id="f39c6-553">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-553">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f39c6-554">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="f39c6-554">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-555">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-555">Type:</span></span>

*   <span data-ttu-id="f39c6-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="f39c6-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-557">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-557">Requirements</span></span>

|<span data-ttu-id="f39c6-558">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-558">Requirement</span></span>| <span data-ttu-id="f39c6-559">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-560">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-560">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-561">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-561">1.0</span></span>|
|[<span data-ttu-id="f39c6-562">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-562">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-563">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-564">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-564">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-565">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-565">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-566">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-566">Example</span></span>

<span data-ttu-id="f39c6-567">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-567">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="f39c6-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f39c6-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="f39c6-569">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f39c6-569">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f39c6-570">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="f39c6-570">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f39c6-571">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-571">Read mode</span></span>

<span data-ttu-id="f39c6-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="f39c6-574">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-574">Compose mode</span></span>

<span data-ttu-id="f39c6-575">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="f39c6-575">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f39c6-576">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-576">Type:</span></span>

*   <span data-ttu-id="f39c6-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f39c6-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-578">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-578">Requirements</span></span>

|<span data-ttu-id="f39c6-579">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-579">Requirement</span></span>| <span data-ttu-id="f39c6-580">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-581">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-581">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-582">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-582">1.0</span></span>|
|[<span data-ttu-id="f39c6-583">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-584">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-585">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-586">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-586">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="f39c6-587">à : Array. <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f39c6-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="f39c6-588">Permet d’accéder aux destinataires de la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="f39c6-588">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f39c6-589">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-589">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f39c6-590">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-590">Read mode</span></span>

<span data-ttu-id="f39c6-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f39c6-593">Mode composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-593">Compose mode</span></span>

<span data-ttu-id="f39c6-594">Le `to` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="f39c6-594">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f39c6-595">Type :</span><span class="sxs-lookup"><span data-stu-id="f39c6-595">Type:</span></span>

*   <span data-ttu-id="f39c6-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f39c6-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-597">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-597">Requirements</span></span>

|<span data-ttu-id="f39c6-598">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-598">Requirement</span></span>| <span data-ttu-id="f39c6-599">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-599">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-600">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-600">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-601">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-601">1.0</span></span>|
|[<span data-ttu-id="f39c6-602">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-602">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-603">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-603">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-604">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-604">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-605">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-605">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-606">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-606">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="f39c6-607">Méthodes</span><span class="sxs-lookup"><span data-stu-id="f39c6-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f39c6-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f39c6-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f39c6-609">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="f39c6-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f39c6-610">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="f39c6-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f39c6-611">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="f39c6-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f39c6-612">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="f39c6-612">Parameters:</span></span>

|<span data-ttu-id="f39c6-613">Nom</span><span class="sxs-lookup"><span data-stu-id="f39c6-613">Name</span></span>| <span data-ttu-id="f39c6-614">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-614">Type</span></span>| <span data-ttu-id="f39c6-615">Attributs</span><span class="sxs-lookup"><span data-stu-id="f39c6-615">Attributes</span></span>| <span data-ttu-id="f39c6-616">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="f39c6-617">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-617">String</span></span>||<span data-ttu-id="f39c6-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f39c6-620">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-620">String</span></span>||<span data-ttu-id="f39c6-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f39c6-623">Objet</span><span class="sxs-lookup"><span data-stu-id="f39c6-623">Object</span></span>| <span data-ttu-id="f39c6-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-624">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-625">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f39c6-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="f39c6-626">Objet</span><span class="sxs-lookup"><span data-stu-id="f39c6-626">Object</span></span> | <span data-ttu-id="f39c6-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-627">&lt;optional&gt;</span></span> | <span data-ttu-id="f39c6-628">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="f39c6-629">Boolean</span><span class="sxs-lookup"><span data-stu-id="f39c6-629">Boolean</span></span> | <span data-ttu-id="f39c6-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-630">&lt;optional&gt;</span></span> | <span data-ttu-id="f39c6-631">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f39c6-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="f39c6-632">fonction</span><span class="sxs-lookup"><span data-stu-id="f39c6-632">function</span></span>| <span data-ttu-id="f39c6-633">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-633">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-634">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f39c6-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f39c6-635">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f39c6-636">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="f39c6-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f39c6-637">Erreurs</span><span class="sxs-lookup"><span data-stu-id="f39c6-637">Errors</span></span>

| <span data-ttu-id="f39c6-638">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="f39c6-638">Error code</span></span> | <span data-ttu-id="f39c6-639">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="f39c6-640">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="f39c6-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="f39c6-641">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="f39c6-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f39c6-642">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f39c6-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f39c6-643">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-643">Requirements</span></span>

|<span data-ttu-id="f39c6-644">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-644">Requirement</span></span>| <span data-ttu-id="f39c6-645">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-646">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-646">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-647">1.1</span><span class="sxs-lookup"><span data-stu-id="f39c6-647">1.1</span></span>|
|[<span data-ttu-id="f39c6-648">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="f39c6-650">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-651">Composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f39c6-652">範例</span><span class="sxs-lookup"><span data-stu-id="f39c6-652">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="f39c6-653">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="f39c6-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f39c6-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f39c6-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f39c6-655">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f39c6-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f39c6-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f39c6-659">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="f39c6-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f39c6-660">Si votre complément Office est en cours d’exécution dans Outlook Web App, le `addItemAttachmentAsync` méthode permettre joindre des éléments à des éléments autres que l’élément à modifier ; Toutefois, cela n’est pas pris en charge et n’est pas recommandée.</span><span class="sxs-lookup"><span data-stu-id="f39c6-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f39c6-661">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="f39c6-661">Parameters:</span></span>

|<span data-ttu-id="f39c6-662">Nom</span><span class="sxs-lookup"><span data-stu-id="f39c6-662">Name</span></span>| <span data-ttu-id="f39c6-663">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-663">Type</span></span>| <span data-ttu-id="f39c6-664">Attributs</span><span class="sxs-lookup"><span data-stu-id="f39c6-664">Attributes</span></span>| <span data-ttu-id="f39c6-665">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="f39c6-666">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-666">String</span></span>||<span data-ttu-id="f39c6-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f39c6-669">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-669">String</span></span>||<span data-ttu-id="f39c6-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f39c6-672">Objet</span><span class="sxs-lookup"><span data-stu-id="f39c6-672">Object</span></span>| <span data-ttu-id="f39c6-673">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-673">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-674">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f39c6-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f39c6-675">Objet</span><span class="sxs-lookup"><span data-stu-id="f39c6-675">Object</span></span>| <span data-ttu-id="f39c6-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-676">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-677">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f39c6-678">fonction</span><span class="sxs-lookup"><span data-stu-id="f39c6-678">function</span></span>| <span data-ttu-id="f39c6-679">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-679">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-680">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f39c6-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f39c6-681">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f39c6-682">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="f39c6-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f39c6-683">Erreurs</span><span class="sxs-lookup"><span data-stu-id="f39c6-683">Errors</span></span>

| <span data-ttu-id="f39c6-684">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="f39c6-684">Error code</span></span> | <span data-ttu-id="f39c6-685">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f39c6-686">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f39c6-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f39c6-687">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-687">Requirements</span></span>

|<span data-ttu-id="f39c6-688">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-688">Requirement</span></span>| <span data-ttu-id="f39c6-689">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-690">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-690">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-691">1.1</span><span class="sxs-lookup"><span data-stu-id="f39c6-691">1.1</span></span>|
|[<span data-ttu-id="f39c6-692">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="f39c6-694">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-695">Composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-696">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-696">Example</span></span>

<span data-ttu-id="f39c6-697">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="f39c6-698">close()</span><span class="sxs-lookup"><span data-stu-id="f39c6-698">close()</span></span>

<span data-ttu-id="f39c6-699">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="f39c6-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="f39c6-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-702">Dans Outlook sur le web, si l’élément est un rendez-vous et il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, supprimer ou annuler même si aucune modification de l’élément depuis le dernier enregistrée.</span><span class="sxs-lookup"><span data-stu-id="f39c6-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="f39c6-703">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="f39c6-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-704">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-704">Requirements</span></span>

|<span data-ttu-id="f39c6-705">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-705">Requirement</span></span>| <span data-ttu-id="f39c6-706">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-707">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-707">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-708">1.3</span><span class="sxs-lookup"><span data-stu-id="f39c6-708">1.3</span></span>|
|[<span data-ttu-id="f39c6-709">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-710">Restricted</span><span class="sxs-lookup"><span data-stu-id="f39c6-710">Restricted</span></span>|
|[<span data-ttu-id="f39c6-711">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-712">Composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-712">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="f39c6-713">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f39c6-713">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="f39c6-714">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f39c6-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-715">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f39c6-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f39c6-716">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="f39c6-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f39c6-717">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="f39c6-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="f39c6-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f39c6-721">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="f39c6-721">Parameters:</span></span>

| <span data-ttu-id="f39c6-722">Nom</span><span class="sxs-lookup"><span data-stu-id="f39c6-722">Name</span></span> | <span data-ttu-id="f39c6-723">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-723">Type</span></span> | <span data-ttu-id="f39c6-724">Attributs</span><span class="sxs-lookup"><span data-stu-id="f39c6-724">Attributes</span></span> | <span data-ttu-id="f39c6-725">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="f39c6-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f39c6-726">String &#124; Object</span></span>| |<span data-ttu-id="f39c6-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f39c6-729">**OU**</span><span class="sxs-lookup"><span data-stu-id="f39c6-729">**OR**</span></span><br/><span data-ttu-id="f39c6-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="f39c6-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f39c6-732">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-732">String</span></span> | <span data-ttu-id="f39c6-733">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-733">&lt;optional&gt;</span></span> | <span data-ttu-id="f39c6-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f39c6-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f39c6-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-737">&lt;optional&gt;</span></span> | <span data-ttu-id="f39c6-738">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="f39c6-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f39c6-739">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-739">String</span></span> | | <span data-ttu-id="f39c6-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f39c6-742">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-742">String</span></span> | | <span data-ttu-id="f39c6-743">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="f39c6-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f39c6-744">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-744">String</span></span> | | <span data-ttu-id="f39c6-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="f39c6-747">Boolean</span><span class="sxs-lookup"><span data-stu-id="f39c6-747">Boolean</span></span> | | <span data-ttu-id="f39c6-p144">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f39c6-750">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-750">String</span></span> | | <span data-ttu-id="f39c6-p145">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f39c6-754">function</span><span class="sxs-lookup"><span data-stu-id="f39c6-754">function</span></span> | <span data-ttu-id="f39c6-755">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-755">&lt;optional&gt;</span></span> | <span data-ttu-id="f39c6-756">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f39c6-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f39c6-757">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-757">Requirements</span></span>

|<span data-ttu-id="f39c6-758">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-758">Requirement</span></span>| <span data-ttu-id="f39c6-759">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-760">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-760">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-761">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-761">1.0</span></span>|
|[<span data-ttu-id="f39c6-762">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-763">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-764">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-765">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f39c6-766">Exemples</span><span class="sxs-lookup"><span data-stu-id="f39c6-766">Examples</span></span>

<span data-ttu-id="f39c6-767">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f39c6-768">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="f39c6-768">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f39c6-769">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="f39c6-769">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f39c6-770">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="f39c6-770">Reply with a body and a file attachment.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="f39c6-771">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f39c6-771">Reply with a body and an item attachment.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="f39c6-772">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="f39c6-773">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f39c6-773">displayReplyForm(formData)</span></span>

<span data-ttu-id="f39c6-774">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f39c6-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-775">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f39c6-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f39c6-776">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="f39c6-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f39c6-777">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="f39c6-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="f39c6-p146">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f39c6-781">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="f39c6-781">Parameters:</span></span>

| <span data-ttu-id="f39c6-782">Nom</span><span class="sxs-lookup"><span data-stu-id="f39c6-782">Name</span></span> | <span data-ttu-id="f39c6-783">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-783">Type</span></span> | <span data-ttu-id="f39c6-784">Attributs</span><span class="sxs-lookup"><span data-stu-id="f39c6-784">Attributes</span></span> | <span data-ttu-id="f39c6-785">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="f39c6-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="f39c6-786">String &#124; Object</span></span>| | <span data-ttu-id="f39c6-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f39c6-789">**OU**</span><span class="sxs-lookup"><span data-stu-id="f39c6-789">**OR**</span></span><br/><span data-ttu-id="f39c6-p148">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="f39c6-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f39c6-792">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-792">String</span></span> | <span data-ttu-id="f39c6-793">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-793">&lt;optional&gt;</span></span> | <span data-ttu-id="f39c6-p149">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f39c6-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f39c6-797">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-797">&lt;optional&gt;</span></span> | <span data-ttu-id="f39c6-798">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="f39c6-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f39c6-799">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-799">String</span></span> | | <span data-ttu-id="f39c6-p150">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f39c6-802">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-802">String</span></span> | | <span data-ttu-id="f39c6-803">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="f39c6-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f39c6-804">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-804">String</span></span> | | <span data-ttu-id="f39c6-p151">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="f39c6-807">Boolean</span><span class="sxs-lookup"><span data-stu-id="f39c6-807">Boolean</span></span> | | <span data-ttu-id="f39c6-p152">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f39c6-810">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-810">String</span></span> | | <span data-ttu-id="f39c6-p153">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f39c6-814">function</span><span class="sxs-lookup"><span data-stu-id="f39c6-814">function</span></span> | <span data-ttu-id="f39c6-815">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-815">&lt;optional&gt;</span></span> | <span data-ttu-id="f39c6-816">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f39c6-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f39c6-817">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-817">Requirements</span></span>

|<span data-ttu-id="f39c6-818">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-818">Requirement</span></span>| <span data-ttu-id="f39c6-819">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-820">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-820">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-821">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-821">1.0</span></span>|
|[<span data-ttu-id="f39c6-822">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-823">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-824">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-825">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f39c6-826">Exemples</span><span class="sxs-lookup"><span data-stu-id="f39c6-826">Examples</span></span>

<span data-ttu-id="f39c6-827">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f39c6-828">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="f39c6-828">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f39c6-829">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="f39c6-829">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f39c6-830">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="f39c6-830">Reply with a body and a file attachment.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="f39c6-831">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="f39c6-831">Reply with a body and an item attachment.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="f39c6-832">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="f39c6-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f39c6-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="f39c6-834">Obtient les entités trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f39c6-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-835">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f39c6-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-836">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-836">Requirements</span></span>

|<span data-ttu-id="f39c6-837">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-837">Requirement</span></span>| <span data-ttu-id="f39c6-838">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-839">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-839">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-840">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-840">1.0</span></span>|
|[<span data-ttu-id="f39c6-841">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-842">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-843">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-844">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f39c6-845">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f39c6-845">Returns:</span></span>

<span data-ttu-id="f39c6-846">Type : [Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f39c6-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f39c6-847">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-847">Example</span></span>

<span data-ttu-id="f39c6-848">L’exemple suivant accède aux entités de contacts dans le corps de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-848">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="f39c6-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f39c6-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f39c6-850">Obtient un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f39c6-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-851">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f39c6-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f39c6-852">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="f39c6-852">Parameters:</span></span>

|<span data-ttu-id="f39c6-853">Nom</span><span class="sxs-lookup"><span data-stu-id="f39c6-853">Name</span></span>| <span data-ttu-id="f39c6-854">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-854">Type</span></span>| <span data-ttu-id="f39c6-855">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="f39c6-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f39c6-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="f39c6-857">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="f39c6-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f39c6-858">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-858">Requirements</span></span>

|<span data-ttu-id="f39c6-859">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-859">Requirement</span></span>| <span data-ttu-id="f39c6-860">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-861">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-861">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-862">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-862">1.0</span></span>|
|[<span data-ttu-id="f39c6-863">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-864">Restricted</span><span class="sxs-lookup"><span data-stu-id="f39c6-864">Restricted</span></span>|
|[<span data-ttu-id="f39c6-865">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-866">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f39c6-867">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f39c6-867">Returns:</span></span>

<span data-ttu-id="f39c6-868">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="f39c6-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f39c6-869">Si aucune entité du type spécifié n’est présentes dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="f39c6-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="f39c6-870">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f39c6-871">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="f39c6-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="f39c6-872">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="f39c6-872">Value of `entityType`</span></span> | <span data-ttu-id="f39c6-873">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="f39c6-873">Type of objects in returned array</span></span> | <span data-ttu-id="f39c6-874">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="f39c6-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="f39c6-875">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-875">String</span></span> | <span data-ttu-id="f39c6-876">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f39c6-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="f39c6-877">Contact</span><span class="sxs-lookup"><span data-stu-id="f39c6-877">Contact</span></span> | <span data-ttu-id="f39c6-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f39c6-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="f39c6-879">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-879">String</span></span> | <span data-ttu-id="f39c6-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f39c6-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="f39c6-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f39c6-881">MeetingSuggestion</span></span> | <span data-ttu-id="f39c6-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f39c6-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="f39c6-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f39c6-883">PhoneNumber</span></span> | <span data-ttu-id="f39c6-884">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f39c6-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="f39c6-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f39c6-885">TaskSuggestion</span></span> | <span data-ttu-id="f39c6-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f39c6-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="f39c6-887">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-887">String</span></span> | <span data-ttu-id="f39c6-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f39c6-888">**Restricted**</span></span> |

<span data-ttu-id="f39c6-889">Type : Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f39c6-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="f39c6-890">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-890">Example</span></span>

<span data-ttu-id="f39c6-891">L’exemple suivant montre comment accéder à un tableau de chaînes qui représentent les adresses postales dans le corps de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="f39c6-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f39c6-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f39c6-893">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="f39c6-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-894">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f39c6-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f39c6-895">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="f39c6-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f39c6-896">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="f39c6-896">Parameters:</span></span>

|<span data-ttu-id="f39c6-897">Nom</span><span class="sxs-lookup"><span data-stu-id="f39c6-897">Name</span></span>| <span data-ttu-id="f39c6-898">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-898">Type</span></span>| <span data-ttu-id="f39c6-899">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f39c6-900">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-900">String</span></span>|<span data-ttu-id="f39c6-901">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="f39c6-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f39c6-902">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-902">Requirements</span></span>

|<span data-ttu-id="f39c6-903">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-903">Requirement</span></span>| <span data-ttu-id="f39c6-904">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-905">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-906">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-906">1.0</span></span>|
|[<span data-ttu-id="f39c6-907">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-908">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-909">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-910">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f39c6-911">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f39c6-911">Returns:</span></span>

<span data-ttu-id="f39c6-p155">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="f39c6-914">Type : Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f39c6-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="f39c6-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f39c6-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f39c6-916">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="f39c6-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-917">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f39c6-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f39c6-p156">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f39c6-921">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="f39c6-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f39c6-922">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f39c6-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f39c6-926">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-926">Requirements</span></span>

|<span data-ttu-id="f39c6-927">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-927">Requirement</span></span>| <span data-ttu-id="f39c6-928">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-929">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-929">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-930">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-930">1.0</span></span>|
|[<span data-ttu-id="f39c6-931">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-932">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-933">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-934">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f39c6-935">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f39c6-935">Returns:</span></span>

<span data-ttu-id="f39c6-p158">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="f39c6-938">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="f39c6-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f39c6-939">Object</span><span class="sxs-lookup"><span data-stu-id="f39c6-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f39c6-940">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-940">Example</span></span>

<span data-ttu-id="f39c6-941">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="f39c6-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f39c6-942">getregexmatchesbyname (Name) → (nullable) {tableau. < chaîne >}</span><span class="sxs-lookup"><span data-stu-id="f39c6-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f39c6-943">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="f39c6-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-944">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="f39c6-944">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f39c6-945">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="f39c6-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f39c6-p159">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f39c6-948">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="f39c6-948">Parameters:</span></span>

|<span data-ttu-id="f39c6-949">Nom</span><span class="sxs-lookup"><span data-stu-id="f39c6-949">Name</span></span>| <span data-ttu-id="f39c6-950">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-950">Type</span></span>| <span data-ttu-id="f39c6-951">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f39c6-952">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-952">String</span></span>|<span data-ttu-id="f39c6-953">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="f39c6-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f39c6-954">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-954">Requirements</span></span>

|<span data-ttu-id="f39c6-955">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-955">Requirement</span></span>| <span data-ttu-id="f39c6-956">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-957">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-957">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-958">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-958">1.0</span></span>|
|[<span data-ttu-id="f39c6-959">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-960">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-961">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-962">Lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f39c6-963">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f39c6-963">Returns:</span></span>

<span data-ttu-id="f39c6-964">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="f39c6-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="f39c6-965">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="f39c6-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f39c6-966">< Chaîne > tableau.</span><span class="sxs-lookup"><span data-stu-id="f39c6-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f39c6-967">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-967">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="f39c6-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="f39c6-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="f39c6-969">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="f39c6-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="f39c6-p160">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f39c6-972">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="f39c6-972">Parameters:</span></span>

|<span data-ttu-id="f39c6-973">Nom</span><span class="sxs-lookup"><span data-stu-id="f39c6-973">Name</span></span>| <span data-ttu-id="f39c6-974">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-974">Type</span></span>| <span data-ttu-id="f39c6-975">Attributs</span><span class="sxs-lookup"><span data-stu-id="f39c6-975">Attributes</span></span>| <span data-ttu-id="f39c6-976">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="f39c6-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f39c6-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="f39c6-p161">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="f39c6-981">Object</span><span class="sxs-lookup"><span data-stu-id="f39c6-981">Object</span></span>| <span data-ttu-id="f39c6-982">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-982">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-983">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f39c6-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f39c6-984">Objet</span><span class="sxs-lookup"><span data-stu-id="f39c6-984">Object</span></span>| <span data-ttu-id="f39c6-985">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-985">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-986">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f39c6-987">fonction</span><span class="sxs-lookup"><span data-stu-id="f39c6-987">function</span></span>||<span data-ttu-id="f39c6-988">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f39c6-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f39c6-989">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="f39c6-990">Pour accéder à la propriété source provenant de la sélection, appelez `asyncResult.value.sourceProperty`, qui sera soit `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-990">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f39c6-991">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-991">Requirements</span></span>

|<span data-ttu-id="f39c6-992">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-992">Requirement</span></span>| <span data-ttu-id="f39c6-993">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-994">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-994">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-995">1.2</span><span class="sxs-lookup"><span data-stu-id="f39c6-995">1.2</span></span>|
|[<span data-ttu-id="f39c6-996">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="f39c6-998">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-999">Composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f39c6-1000">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="f39c6-1000">Returns:</span></span>

<span data-ttu-id="f39c6-1001">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="f39c6-1002">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="f39c6-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f39c6-1003">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f39c6-1004">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-1004">Example</span></span>

```
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f39c6-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f39c6-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f39c6-1006">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f39c6-p163">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f39c6-1010">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="f39c6-1010">Parameters:</span></span>

|<span data-ttu-id="f39c6-1011">Nom</span><span class="sxs-lookup"><span data-stu-id="f39c6-1011">Name</span></span>| <span data-ttu-id="f39c6-1012">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-1012">Type</span></span>| <span data-ttu-id="f39c6-1013">Attributs</span><span class="sxs-lookup"><span data-stu-id="f39c6-1013">Attributes</span></span>| <span data-ttu-id="f39c6-1014">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f39c6-1015">function</span><span class="sxs-lookup"><span data-stu-id="f39c6-1015">function</span></span>||<span data-ttu-id="f39c6-1016">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f39c6-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f39c6-1017">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f39c6-1018">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées de l’élément et enregistrer les modifications à la propriété personnalisée est redéfinie sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1018">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="f39c6-1019">Objet</span><span class="sxs-lookup"><span data-stu-id="f39c6-1019">Object</span></span>| <span data-ttu-id="f39c6-1020">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-1021">Les développeurs peuvent fournir n’importe quel objet qu’ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1021">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="f39c6-1022">Cet objet est accessible par le `asyncResult.asyncContext` propriété dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f39c6-1023">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-1023">Requirements</span></span>

|<span data-ttu-id="f39c6-1024">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-1024">Requirement</span></span>| <span data-ttu-id="f39c6-1025">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-1026">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-1026">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="f39c6-1027">1.0</span></span>|
|[<span data-ttu-id="f39c6-1028">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-1029">ReadItem</span></span>|
|[<span data-ttu-id="f39c6-1030">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-1031">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="f39c6-1031">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-1032">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-1032">Example</span></span>

<span data-ttu-id="f39c6-p166">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f39c6-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f39c6-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f39c6-1037">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f39c6-p167">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f39c6-1042">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="f39c6-1042">Parameters:</span></span>

|<span data-ttu-id="f39c6-1043">Nom</span><span class="sxs-lookup"><span data-stu-id="f39c6-1043">Name</span></span>| <span data-ttu-id="f39c6-1044">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-1044">Type</span></span>| <span data-ttu-id="f39c6-1045">Attributs</span><span class="sxs-lookup"><span data-stu-id="f39c6-1045">Attributes</span></span>| <span data-ttu-id="f39c6-1046">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="f39c6-1047">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-1047">String</span></span>||<span data-ttu-id="f39c6-p168">Identificateur de la pièce jointe à supprimer. La longueur maximale de la chaîne est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p168">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="f39c6-1050">Objet</span><span class="sxs-lookup"><span data-stu-id="f39c6-1050">Object</span></span>| <span data-ttu-id="f39c6-1051">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-1052">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f39c6-1053">Objet</span><span class="sxs-lookup"><span data-stu-id="f39c6-1053">Object</span></span>| <span data-ttu-id="f39c6-1054">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-1055">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f39c6-1056">fonction</span><span class="sxs-lookup"><span data-stu-id="f39c6-1056">function</span></span>| <span data-ttu-id="f39c6-1057">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-1058">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f39c6-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f39c6-1059">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f39c6-1060">Erreurs</span><span class="sxs-lookup"><span data-stu-id="f39c6-1060">Errors</span></span>

| <span data-ttu-id="f39c6-1061">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="f39c6-1061">Error code</span></span> | <span data-ttu-id="f39c6-1062">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="f39c6-1063">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f39c6-1064">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-1064">Requirements</span></span>

|<span data-ttu-id="f39c6-1065">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-1065">Requirement</span></span>| <span data-ttu-id="f39c6-1066">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-1067">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-1067">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="f39c6-1068">1.1</span></span>|
|[<span data-ttu-id="f39c6-1069">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="f39c6-1071">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-1072">Composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-1073">Exemple</span><span class="sxs-lookup"><span data-stu-id="f39c6-1073">Example</span></span>

<span data-ttu-id="f39c6-1074">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="f39c6-1074">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="f39c6-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="f39c6-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="f39c6-1076">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="f39c6-p169">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p169">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-1080">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` pour utiliser avec EWS ou l’API REST, gardez à l’esprit que quand Outlook est en mode mis en cache, il peut prendre un certain temps avant que l’élément est réellement synchronisé avec le serveur.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1080">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="f39c6-1081">Jusqu'à ce que l’élément est synchronisé, à l’aide de la `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="f39c6-p171">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p171">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="f39c6-1085">Les clients suivants ont un comportement différent pour `saveAsync` de rendez-vous dans le mode de composition :</span><span class="sxs-lookup"><span data-stu-id="f39c6-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="f39c6-1086">Mac Outlook ne gère pas `saveAsync` d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1086">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="f39c6-1087">Appel de `saveAsync` d’une réunion dans Outlook Mac renverra une erreur.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1087">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="f39c6-1088">Outlook sur le web toujours envoie une invitation ou mettre à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f39c6-1089">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="f39c6-1089">Parameters:</span></span>

|<span data-ttu-id="f39c6-1090">Nom</span><span class="sxs-lookup"><span data-stu-id="f39c6-1090">Name</span></span>| <span data-ttu-id="f39c6-1091">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-1091">Type</span></span>| <span data-ttu-id="f39c6-1092">Attributs</span><span class="sxs-lookup"><span data-stu-id="f39c6-1092">Attributes</span></span>| <span data-ttu-id="f39c6-1093">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="f39c6-1094">Objet</span><span class="sxs-lookup"><span data-stu-id="f39c6-1094">Object</span></span>| <span data-ttu-id="f39c6-1095">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-1096">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f39c6-1097">Objet</span><span class="sxs-lookup"><span data-stu-id="f39c6-1097">Object</span></span>| <span data-ttu-id="f39c6-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-1099">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f39c6-1100">fonction</span><span class="sxs-lookup"><span data-stu-id="f39c6-1100">function</span></span>||<span data-ttu-id="f39c6-1101">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f39c6-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f39c6-1102">En cas de réussite, l’identificateur d’élément est fournie dans le `asyncResult.value` propriété.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f39c6-1103">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-1103">Requirements</span></span>

|<span data-ttu-id="f39c6-1104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-1104">Requirement</span></span>| <span data-ttu-id="f39c6-1105">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-1106">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-1106">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="f39c6-1107">1.3</span></span>|
|[<span data-ttu-id="f39c6-1108">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="f39c6-1110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-1111">Composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f39c6-1112">範例</span><span class="sxs-lookup"><span data-stu-id="f39c6-1112">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="f39c6-p173">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p173">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="f39c6-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="f39c6-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="f39c6-1116">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="f39c6-p174">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f39c6-1120">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="f39c6-1120">Parameters:</span></span>

|<span data-ttu-id="f39c6-1121">Nom</span><span class="sxs-lookup"><span data-stu-id="f39c6-1121">Name</span></span>| <span data-ttu-id="f39c6-1122">Type</span><span class="sxs-lookup"><span data-stu-id="f39c6-1122">Type</span></span>| <span data-ttu-id="f39c6-1123">Attributs</span><span class="sxs-lookup"><span data-stu-id="f39c6-1123">Attributes</span></span>| <span data-ttu-id="f39c6-1124">Description</span><span class="sxs-lookup"><span data-stu-id="f39c6-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="f39c6-1125">String</span><span class="sxs-lookup"><span data-stu-id="f39c6-1125">String</span></span>||<span data-ttu-id="f39c6-p175">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="f39c6-1129">Objet</span><span class="sxs-lookup"><span data-stu-id="f39c6-1129">Object</span></span>| <span data-ttu-id="f39c6-1130">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-1131">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f39c6-1132">Objet</span><span class="sxs-lookup"><span data-stu-id="f39c6-1132">Object</span></span>| <span data-ttu-id="f39c6-1133">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-1134">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="f39c6-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f39c6-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="f39c6-1136">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f39c6-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="f39c6-p176">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p176">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="f39c6-p177">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="f39c6-p177">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="f39c6-1141">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="f39c6-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="f39c6-1142">fonction</span><span class="sxs-lookup"><span data-stu-id="f39c6-1142">function</span></span>||<span data-ttu-id="f39c6-1143">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f39c6-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f39c6-1144">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f39c6-1144">Requirements</span></span>

|<span data-ttu-id="f39c6-1145">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f39c6-1145">Requirement</span></span>| <span data-ttu-id="f39c6-1146">Valeur</span><span class="sxs-lookup"><span data-stu-id="f39c6-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="f39c6-1147">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f39c6-1147">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f39c6-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="f39c6-1148">1.2</span></span>|
|[<span data-ttu-id="f39c6-1149">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="f39c6-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f39c6-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f39c6-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="f39c6-1151">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f39c6-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f39c6-1152">Composition</span><span class="sxs-lookup"><span data-stu-id="f39c6-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f39c6-1153">Exemples</span><span class="sxs-lookup"><span data-stu-id="f39c6-1153">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```