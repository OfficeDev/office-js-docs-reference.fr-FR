
# <a name="item"></a><span data-ttu-id="6cf1a-101">élément</span><span class="sxs-lookup"><span data-stu-id="6cf1a-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="6cf1a-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="6cf1a-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="6cf1a-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-105">Requirements</span></span>

|<span data-ttu-id="6cf1a-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-106">Requirement</span></span>|<span data-ttu-id="6cf1a-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-109">1.0</span></span>|
|[<span data-ttu-id="6cf1a-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-111">Restricted</span><span class="sxs-lookup"><span data-stu-id="6cf1a-111">Restricted</span></span>|
|[<span data-ttu-id="6cf1a-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6cf1a-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="6cf1a-114">Members and methods</span></span>

| <span data-ttu-id="6cf1a-115">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-115">Member</span></span> | <span data-ttu-id="6cf1a-116">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6cf1a-117">attachments</span><span class="sxs-lookup"><span data-stu-id="6cf1a-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="6cf1a-118">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-118">Member</span></span> |
| [<span data-ttu-id="6cf1a-119">bcc</span><span class="sxs-lookup"><span data-stu-id="6cf1a-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="6cf1a-120">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-120">Member</span></span> |
| [<span data-ttu-id="6cf1a-121">body</span><span class="sxs-lookup"><span data-stu-id="6cf1a-121">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="6cf1a-122">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-122">Member</span></span> |
| [<span data-ttu-id="6cf1a-123">cc</span><span class="sxs-lookup"><span data-stu-id="6cf1a-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="6cf1a-124">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-124">Member</span></span> |
| [<span data-ttu-id="6cf1a-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="6cf1a-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="6cf1a-126">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-126">Member</span></span> |
| [<span data-ttu-id="6cf1a-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="6cf1a-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="6cf1a-128">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-128">Member</span></span> |
| [<span data-ttu-id="6cf1a-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="6cf1a-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="6cf1a-130">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-130">Member</span></span> |
| [<span data-ttu-id="6cf1a-131">end</span><span class="sxs-lookup"><span data-stu-id="6cf1a-131">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="6cf1a-132">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-132">Member</span></span> |
| [<span data-ttu-id="6cf1a-133">from</span><span class="sxs-lookup"><span data-stu-id="6cf1a-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="6cf1a-134">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-134">Member</span></span> |
| [<span data-ttu-id="6cf1a-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="6cf1a-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="6cf1a-136">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-136">Member</span></span> |
| [<span data-ttu-id="6cf1a-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="6cf1a-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="6cf1a-138">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-138">Member</span></span> |
| [<span data-ttu-id="6cf1a-139">itemId</span><span class="sxs-lookup"><span data-stu-id="6cf1a-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="6cf1a-140">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-140">Member</span></span> |
| [<span data-ttu-id="6cf1a-141">itemType</span><span class="sxs-lookup"><span data-stu-id="6cf1a-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="6cf1a-142">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-142">Member</span></span> |
| [<span data-ttu-id="6cf1a-143">location</span><span class="sxs-lookup"><span data-stu-id="6cf1a-143">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="6cf1a-144">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-144">Member</span></span> |
| [<span data-ttu-id="6cf1a-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="6cf1a-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="6cf1a-146">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-146">Member</span></span> |
| [<span data-ttu-id="6cf1a-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="6cf1a-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="6cf1a-148">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-148">Member</span></span> |
| [<span data-ttu-id="6cf1a-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="6cf1a-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="6cf1a-150">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-150">Member</span></span> |
| [<span data-ttu-id="6cf1a-151">organizer</span><span class="sxs-lookup"><span data-stu-id="6cf1a-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="6cf1a-152">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-152">Member</span></span> |
| [<span data-ttu-id="6cf1a-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="6cf1a-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="6cf1a-154">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-154">Member</span></span> |
| [<span data-ttu-id="6cf1a-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="6cf1a-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="6cf1a-156">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-156">Member</span></span> |
| [<span data-ttu-id="6cf1a-157">sender</span><span class="sxs-lookup"><span data-stu-id="6cf1a-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="6cf1a-158">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-158">Member</span></span> |
| [<span data-ttu-id="6cf1a-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="6cf1a-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="6cf1a-160">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-160">Member</span></span> |
| [<span data-ttu-id="6cf1a-161">start</span><span class="sxs-lookup"><span data-stu-id="6cf1a-161">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="6cf1a-162">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-162">Member</span></span> |
| [<span data-ttu-id="6cf1a-163">subject</span><span class="sxs-lookup"><span data-stu-id="6cf1a-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="6cf1a-164">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-164">Member</span></span> |
| [<span data-ttu-id="6cf1a-165">to</span><span class="sxs-lookup"><span data-stu-id="6cf1a-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="6cf1a-166">Membre</span><span class="sxs-lookup"><span data-stu-id="6cf1a-166">Member</span></span> |
| [<span data-ttu-id="6cf1a-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="6cf1a-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="6cf1a-168">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-168">Method</span></span> |
| [<span data-ttu-id="6cf1a-169">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="6cf1a-169">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="6cf1a-170">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-170">Method</span></span> |
| [<span data-ttu-id="6cf1a-171">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="6cf1a-171">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="6cf1a-172">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-172">Method</span></span> |
| [<span data-ttu-id="6cf1a-173">close</span><span class="sxs-lookup"><span data-stu-id="6cf1a-173">close</span></span>](#close) | <span data-ttu-id="6cf1a-174">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-174">Method</span></span> |
| [<span data-ttu-id="6cf1a-175">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="6cf1a-175">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="6cf1a-176">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-176">Method</span></span> |
| [<span data-ttu-id="6cf1a-177">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="6cf1a-177">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="6cf1a-178">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-178">Method</span></span> |
| [<span data-ttu-id="6cf1a-179">getEntities</span><span class="sxs-lookup"><span data-stu-id="6cf1a-179">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="6cf1a-180">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-180">Method</span></span> |
| [<span data-ttu-id="6cf1a-181">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="6cf1a-181">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="6cf1a-182">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-182">Method</span></span> |
| [<span data-ttu-id="6cf1a-183">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="6cf1a-183">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="6cf1a-184">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-184">Method</span></span> |
| [<span data-ttu-id="6cf1a-185">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="6cf1a-185">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="6cf1a-186">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-186">Method</span></span> |
| [<span data-ttu-id="6cf1a-187">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="6cf1a-187">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="6cf1a-188">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-188">Method</span></span> |
| [<span data-ttu-id="6cf1a-189">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="6cf1a-189">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="6cf1a-190">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-190">Method</span></span> |
| [<span data-ttu-id="6cf1a-191">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="6cf1a-191">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="6cf1a-192">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-192">Method</span></span> |
| [<span data-ttu-id="6cf1a-193">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="6cf1a-193">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="6cf1a-194">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-194">Method</span></span> |
| [<span data-ttu-id="6cf1a-195">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="6cf1a-195">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="6cf1a-196">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-196">Method</span></span> |
| [<span data-ttu-id="6cf1a-197">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="6cf1a-197">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="6cf1a-198">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-198">Method</span></span> |
| [<span data-ttu-id="6cf1a-199">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="6cf1a-199">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="6cf1a-200">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-200">Method</span></span> |
| [<span data-ttu-id="6cf1a-201">saveAsync</span><span class="sxs-lookup"><span data-stu-id="6cf1a-201">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="6cf1a-202">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-202">Method</span></span> |
| [<span data-ttu-id="6cf1a-203">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="6cf1a-203">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="6cf1a-204">Méthode</span><span class="sxs-lookup"><span data-stu-id="6cf1a-204">Method</span></span> |

### <a name="example"></a><span data-ttu-id="6cf1a-205">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-205">Example</span></span>

<span data-ttu-id="6cf1a-206">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-206">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="6cf1a-207">Membres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-207">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="6cf1a-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="6cf1a-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="6cf1a-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-211">Certains types de fichiers bloqués par Outlook en raison de problèmes de sécurité potentiels et ne sont donc pas retournés.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-211">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="6cf1a-212">Pour plus d’informations, voir les [pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-212">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-213">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-213">Type:</span></span>

*   <span data-ttu-id="6cf1a-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="6cf1a-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-215">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-215">Requirements</span></span>

|<span data-ttu-id="6cf1a-216">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-216">Requirement</span></span>|<span data-ttu-id="6cf1a-217">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-218">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-218">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-219">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-219">1.0</span></span>|
|[<span data-ttu-id="6cf1a-220">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-221">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-222">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-223">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-223">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-224">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-224">Example</span></span>

<span data-ttu-id="6cf1a-225">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-225">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="6cf1a-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="6cf1a-227">Obtient un objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires des Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-227">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="6cf1a-228">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-228">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-229">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-229">Type:</span></span>

*   [<span data-ttu-id="6cf1a-230">Destinataires</span><span class="sxs-lookup"><span data-stu-id="6cf1a-230">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="6cf1a-231">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-231">Requirements</span></span>

|<span data-ttu-id="6cf1a-232">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-232">Requirement</span></span>|<span data-ttu-id="6cf1a-233">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-234">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-234">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-235">1.1</span><span class="sxs-lookup"><span data-stu-id="6cf1a-235">1.1</span></span>|
|[<span data-ttu-id="6cf1a-236">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-236">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-237">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-237">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-238">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-238">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-239">Composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-239">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-240">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-240">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="6cf1a-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="6cf1a-242">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-242">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-243">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-243">Type:</span></span>

*   [<span data-ttu-id="6cf1a-244">Corps</span><span class="sxs-lookup"><span data-stu-id="6cf1a-244">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="6cf1a-245">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-245">Requirements</span></span>

|<span data-ttu-id="6cf1a-246">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-246">Requirement</span></span>|<span data-ttu-id="6cf1a-247">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-248">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-248">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-249">1.1</span><span class="sxs-lookup"><span data-stu-id="6cf1a-249">1.1</span></span>|
|[<span data-ttu-id="6cf1a-250">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-250">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-251">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-252">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-252">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-253">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-253">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="6cf1a-254">cc : tableau. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-254">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="6cf1a-255">Permet d’accéder aux destinataires Cc (copie carbone) d’un message.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="6cf1a-256">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cf1a-257">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-257">Read mode</span></span>

<span data-ttu-id="6cf1a-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cf1a-260">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-260">Compose mode</span></span>

<span data-ttu-id="6cf1a-261">Le `cc` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-261">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-262">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-262">Type:</span></span>

*   <span data-ttu-id="6cf1a-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-264">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-264">Requirements</span></span>

|<span data-ttu-id="6cf1a-265">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-265">Requirement</span></span>|<span data-ttu-id="6cf1a-266">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-267">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-267">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-268">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-268">1.0</span></span>|
|[<span data-ttu-id="6cf1a-269">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-270">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-271">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-272">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-272">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-273">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-273">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="6cf1a-274">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-274">(nullable) conversationId :String</span></span>

<span data-ttu-id="6cf1a-275">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="6cf1a-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="6cf1a-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-280">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-280">Type:</span></span>

*   <span data-ttu-id="6cf1a-281">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6cf1a-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-282">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-282">Requirements</span></span>

|<span data-ttu-id="6cf1a-283">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-283">Requirement</span></span>|<span data-ttu-id="6cf1a-284">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-285">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-285">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-286">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-286">1.0</span></span>|
|[<span data-ttu-id="6cf1a-287">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-288">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-289">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-290">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-290">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="6cf1a-291">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="6cf1a-291">dateTimeCreated :Date</span></span>

<span data-ttu-id="6cf1a-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-294">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-294">Type:</span></span>

*   <span data-ttu-id="6cf1a-295">Date</span><span class="sxs-lookup"><span data-stu-id="6cf1a-295">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-296">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-296">Requirements</span></span>

|<span data-ttu-id="6cf1a-297">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-297">Requirement</span></span>|<span data-ttu-id="6cf1a-298">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-298">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-299">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-299">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-300">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-300">1.0</span></span>|
|[<span data-ttu-id="6cf1a-301">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-301">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-302">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-303">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-303">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-304">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-304">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-305">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-305">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="6cf1a-306">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="6cf1a-306">dateTimeModified :Date</span></span>

<span data-ttu-id="6cf1a-p110">Obtient la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-309">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-309">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-310">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-310">Type:</span></span>

*   <span data-ttu-id="6cf1a-311">Date</span><span class="sxs-lookup"><span data-stu-id="6cf1a-311">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-312">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-312">Requirements</span></span>

|<span data-ttu-id="6cf1a-313">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-313">Requirement</span></span>|<span data-ttu-id="6cf1a-314">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-314">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-315">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-315">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-316">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-316">1.0</span></span>|
|[<span data-ttu-id="6cf1a-317">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-318">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-318">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-319">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-320">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-320">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-321">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-321">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="6cf1a-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="6cf1a-323">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-323">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="6cf1a-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cf1a-326">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-326">Read mode</span></span>

<span data-ttu-id="6cf1a-327">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-327">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cf1a-328">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-328">Compose mode</span></span>

<span data-ttu-id="6cf1a-329">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-329">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="6cf1a-330">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-330">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-331">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-331">Type:</span></span>

*   <span data-ttu-id="6cf1a-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-333">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-333">Requirements</span></span>

|<span data-ttu-id="6cf1a-334">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-334">Requirement</span></span>|<span data-ttu-id="6cf1a-335">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-336">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-337">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-337">1.0</span></span>|
|[<span data-ttu-id="6cf1a-338">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-339">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-340">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-341">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-342">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-342">Example</span></span>

<span data-ttu-id="6cf1a-343">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-343">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="6cf1a-344">à partir de :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[à partir de](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="6cf1a-345">Obtient l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-345">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="6cf1a-p112">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-348">Le `recipientType` propriété de la `EmailAddressDetails` objet dans le `from` propriété est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-348">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cf1a-349">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-349">Read mode</span></span>

<span data-ttu-id="6cf1a-350">Le `from` propriété retourne un `EmailAddressDetails` objet.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-350">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="6cf1a-351">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-351">Compose mode</span></span>

<span data-ttu-id="6cf1a-352">Le `from` propriété renvoie une `From` objet qui fournit une méthode pour obtenir la valeur affectée à.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-352">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="6cf1a-353">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-353">Type:</span></span>

*   <span data-ttu-id="6cf1a-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [à partir de](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-355">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-355">Requirements</span></span>

|<span data-ttu-id="6cf1a-356">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-356">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="6cf1a-357">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-357">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-358">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-358">1.0</span></span>|<span data-ttu-id="6cf1a-359">1.7</span><span class="sxs-lookup"><span data-stu-id="6cf1a-359">1.7</span></span>|
|[<span data-ttu-id="6cf1a-360">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-361">ReadItem</span></span>|<span data-ttu-id="6cf1a-362">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-362">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cf1a-363">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-364">Read</span><span class="sxs-lookup"><span data-stu-id="6cf1a-364">Read</span></span>|<span data-ttu-id="6cf1a-365">Composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-365">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="6cf1a-366">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-366">internetMessageId :String</span></span>

<span data-ttu-id="6cf1a-p113">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-369">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-369">Type:</span></span>

*   <span data-ttu-id="6cf1a-370">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6cf1a-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-371">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-371">Requirements</span></span>

|<span data-ttu-id="6cf1a-372">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-372">Requirement</span></span>|<span data-ttu-id="6cf1a-373">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-374">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-374">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-375">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-375">1.0</span></span>|
|[<span data-ttu-id="6cf1a-376">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-376">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-377">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-378">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-378">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-379">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-380">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-380">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="6cf1a-381">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-381">itemClass :String</span></span>

<span data-ttu-id="6cf1a-p114">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="6cf1a-p115">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="6cf1a-386">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-386">Type</span></span>|<span data-ttu-id="6cf1a-387">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-387">Description</span></span>|<span data-ttu-id="6cf1a-388">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="6cf1a-388">item class</span></span>|
|---|---|---|
|<span data-ttu-id="6cf1a-389">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="6cf1a-389">Appointment items</span></span>|<span data-ttu-id="6cf1a-390">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="6cf1a-391">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="6cf1a-391">Message items</span></span>|<span data-ttu-id="6cf1a-392">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="6cf1a-393">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-394">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-394">Type:</span></span>

*   <span data-ttu-id="6cf1a-395">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6cf1a-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-396">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-396">Requirements</span></span>

|<span data-ttu-id="6cf1a-397">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-397">Requirement</span></span>|<span data-ttu-id="6cf1a-398">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-399">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-399">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-400">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-400">1.0</span></span>|
|[<span data-ttu-id="6cf1a-401">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-402">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-403">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-404">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-405">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-405">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="6cf1a-406">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-406">(nullable) itemId :String</span></span>

<span data-ttu-id="6cf1a-p116">Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-409">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-409">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="6cf1a-410">Le `itemId` propriété n’est pas identique à l’identificateur d’entrée Outlook ou l’ID utilisé par l’API REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="6cf1a-411">Avant d’effectuer des appels d’API REST à l’aide de cette valeur, elle doit être convertie à l’aide de [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="6cf1a-412">Pour plus d’informations, voir [utiliser les API REST d’Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="6cf1a-p118">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-415">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-415">Type:</span></span>

*   <span data-ttu-id="6cf1a-416">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6cf1a-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-417">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-417">Requirements</span></span>

|<span data-ttu-id="6cf1a-418">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-418">Requirement</span></span>|<span data-ttu-id="6cf1a-419">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-420">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-420">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-421">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-421">1.0</span></span>|
|[<span data-ttu-id="6cf1a-422">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-423">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-424">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-425">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-426">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-426">Example</span></span>

<span data-ttu-id="6cf1a-p119">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="6cf1a-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="6cf1a-430">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="6cf1a-431">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-432">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-432">Type:</span></span>

*   [<span data-ttu-id="6cf1a-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="6cf1a-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="6cf1a-434">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-434">Requirements</span></span>

|<span data-ttu-id="6cf1a-435">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-435">Requirement</span></span>|<span data-ttu-id="6cf1a-436">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-437">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-437">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-438">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-438">1.0</span></span>|
|[<span data-ttu-id="6cf1a-439">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-439">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-440">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-441">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-441">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-442">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-442">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-443">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-443">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="6cf1a-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="6cf1a-445">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cf1a-446">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-446">Read mode</span></span>

<span data-ttu-id="6cf1a-447">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-447">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cf1a-448">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-448">Compose mode</span></span>

<span data-ttu-id="6cf1a-449">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-450">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-450">Type:</span></span>

*   <span data-ttu-id="6cf1a-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-452">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-452">Requirements</span></span>

|<span data-ttu-id="6cf1a-453">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-453">Requirement</span></span>|<span data-ttu-id="6cf1a-454">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-455">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-455">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-456">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-456">1.0</span></span>|
|[<span data-ttu-id="6cf1a-457">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-458">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-459">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-460">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-460">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-461">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-461">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="6cf1a-462">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-462">normalizedSubject :String</span></span>

<span data-ttu-id="6cf1a-p120">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="6cf1a-p121">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-467">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-467">Type:</span></span>

*   <span data-ttu-id="6cf1a-468">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6cf1a-468">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-469">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-469">Requirements</span></span>

|<span data-ttu-id="6cf1a-470">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-470">Requirement</span></span>|<span data-ttu-id="6cf1a-471">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-472">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-472">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-473">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-473">1.0</span></span>|
|[<span data-ttu-id="6cf1a-474">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-474">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-475">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-476">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-476">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-477">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-477">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-478">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-478">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="6cf1a-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="6cf1a-480">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-480">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-481">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-481">Type:</span></span>

*   [<span data-ttu-id="6cf1a-482">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="6cf1a-482">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="6cf1a-483">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-483">Requirements</span></span>

|<span data-ttu-id="6cf1a-484">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-484">Requirement</span></span>|<span data-ttu-id="6cf1a-485">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-486">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-486">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-487">1.3</span><span class="sxs-lookup"><span data-stu-id="6cf1a-487">1.3</span></span>|
|[<span data-ttu-id="6cf1a-488">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-488">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-489">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-490">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-490">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-491">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-491">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="6cf1a-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="6cf1a-493">Fournit l’accès à un événement les participants facultatifs.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="6cf1a-494">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cf1a-495">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-495">Read mode</span></span>

<span data-ttu-id="6cf1a-496">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cf1a-497">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-497">Compose mode</span></span>

<span data-ttu-id="6cf1a-498">Le `optionalAttendees` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs à une réunion.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-498">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-499">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-499">Type:</span></span>

*   <span data-ttu-id="6cf1a-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-501">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-501">Requirements</span></span>

|<span data-ttu-id="6cf1a-502">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-502">Requirement</span></span>|<span data-ttu-id="6cf1a-503">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-503">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-504">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-504">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-505">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-505">1.0</span></span>|
|[<span data-ttu-id="6cf1a-506">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-506">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-507">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-507">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-508">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-508">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-509">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-509">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-510">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-510">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="6cf1a-511">organisateur :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[organisateur](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="6cf1a-512">Obtient l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-512">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cf1a-513">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-513">Read mode</span></span>

<span data-ttu-id="6cf1a-514">Le `organizer` propriété renvoie un objet [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-514">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cf1a-515">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-515">Compose mode</span></span>

<span data-ttu-id="6cf1a-516">Le `organizer` propriété retourne un objet [organisateur](/javascript/api/outlook_1_7/office.organizer) qui fournit une méthode pour obtenir la valeur de l’organisateur.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-516">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-517">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-517">Type:</span></span>

*   <span data-ttu-id="6cf1a-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [organisateur](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-519">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-519">Requirements</span></span>

|<span data-ttu-id="6cf1a-520">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-520">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="6cf1a-521">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-522">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-522">1.0</span></span>|<span data-ttu-id="6cf1a-523">1.7</span><span class="sxs-lookup"><span data-stu-id="6cf1a-523">1.7</span></span>|
|[<span data-ttu-id="6cf1a-524">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-525">ReadItem</span></span>|<span data-ttu-id="6cf1a-526">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-526">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cf1a-527">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-527">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-528">Read</span><span class="sxs-lookup"><span data-stu-id="6cf1a-528">Read</span></span>|<span data-ttu-id="6cf1a-529">Composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-530">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-530">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="6cf1a-531">périodicité (nullable) :[périodicité](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="6cf1a-532">Obtient ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-532">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="6cf1a-533">Obtient la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-533">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="6cf1a-534">Lecture et de composition modes pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-534">Read and compose modes for appointment items.</span></span> <span data-ttu-id="6cf1a-535">Mode de lecture pour les éléments de la demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-535">Read mode for meeting request items.</span></span>

<span data-ttu-id="6cf1a-536">Le `recurrence` propriété renvoie un objet de [périodicité](/javascript/api/outlook_1_7/office.recurrence) pour les demandes de réunions ou de rendez-vous périodiques si un élément est une série ou à une instance d’une série.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-536">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="6cf1a-537">`null`sont retournées pour les rendez-vous uniques et des demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-537">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="6cf1a-538">`undefined`sont retournées pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-538">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="6cf1a-539">Remarque : Les demandes de réunion ont une `itemClass` valeur IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-539">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="6cf1a-540">Remarque : Si l’objet de périodicité est `null`, cela indique que l’objet est un rendez-vous ou une demande de réunion d’un rendez-vous et ne fait pas partie d’une série.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-540">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-541">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-541">Type:</span></span>

* [<span data-ttu-id="6cf1a-542">Périodicité</span><span class="sxs-lookup"><span data-stu-id="6cf1a-542">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="6cf1a-543">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-543">Requirement</span></span>|<span data-ttu-id="6cf1a-544">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-545">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-545">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-546">1.7</span><span class="sxs-lookup"><span data-stu-id="6cf1a-546">1.7</span></span>|
|[<span data-ttu-id="6cf1a-547">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-547">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-548">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-549">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-549">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-550">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-550">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="6cf1a-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="6cf1a-552">Fournit l’accès à des participants obligatoires d’un événement.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-552">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="6cf1a-553">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-553">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cf1a-554">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-554">Read mode</span></span>

<span data-ttu-id="6cf1a-555">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-555">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cf1a-556">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-556">Compose mode</span></span>

<span data-ttu-id="6cf1a-557">Le `requiredAttendees` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les participants obligatoires à une réunion.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-557">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-558">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-558">Type:</span></span>

*   <span data-ttu-id="6cf1a-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-560">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-560">Requirements</span></span>

|<span data-ttu-id="6cf1a-561">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-561">Requirement</span></span>|<span data-ttu-id="6cf1a-562">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-563">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-563">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-564">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-564">1.0</span></span>|
|[<span data-ttu-id="6cf1a-565">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-566">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-567">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-568">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-568">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-569">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-569">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="6cf1a-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="6cf1a-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="6cf1a-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-575">Le `recipientType` propriété de la `EmailAddressDetails` objet dans le `sender` propriété est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-575">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-576">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-576">Type:</span></span>

*   [<span data-ttu-id="6cf1a-577">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6cf1a-577">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="6cf1a-578">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-578">Requirements</span></span>

|<span data-ttu-id="6cf1a-579">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-579">Requirement</span></span>|<span data-ttu-id="6cf1a-580">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-581">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-581">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-582">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-582">1.0</span></span>|
|[<span data-ttu-id="6cf1a-583">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-584">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-585">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-586">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-586">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-587">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-587">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="6cf1a-588">seriesId (nullable) : chaîne</span><span class="sxs-lookup"><span data-stu-id="6cf1a-588">(nullable) seriesId :String</span></span>

<span data-ttu-id="6cf1a-589">Obtient l’id de la série appartenant à une instance.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-589">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="6cf1a-590">Dans OWA et Outlook, le `seriesId` renvoie l’ID Exchange Web Services (EWS) de l’élément parent (série) auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-590">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="6cf1a-591">Toutefois, dans iOS et Android, le `seriesId` renvoie l’ID du reste de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-591">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-592">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-592">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="6cf1a-593">Le `seriesId` propriété n’est pas identique à l’ID Outlook utilisées par l’API REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-593">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="6cf1a-594">Avant d’effectuer des appels d’API REST à l’aide de cette valeur, elle doit être convertie à l’aide de [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-594">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="6cf1a-595">Pour plus d’informations, voir [utiliser les API REST d’Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-595">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="6cf1a-596">Le `seriesId` propriété renvoie `null` pour les éléments qui n’ont pas d’éléments parents comme unique rendez-vous, les éléments de série, ou les demandes de réunion et renvoie `undefined` pour tous les éléments qui ne sont pas les demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-596">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-597">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-597">Type:</span></span>

* <span data-ttu-id="6cf1a-598">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6cf1a-598">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-599">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-599">Requirements</span></span>

|<span data-ttu-id="6cf1a-600">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-600">Requirement</span></span>|<span data-ttu-id="6cf1a-601">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-602">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-602">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-603">1.7</span><span class="sxs-lookup"><span data-stu-id="6cf1a-603">1.7</span></span>|
|[<span data-ttu-id="6cf1a-604">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-604">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-605">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-606">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-606">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-607">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-607">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-608">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-608">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="6cf1a-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="6cf1a-610">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-610">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="6cf1a-p130">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cf1a-613">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-613">Read mode</span></span>

<span data-ttu-id="6cf1a-614">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-614">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cf1a-615">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-615">Compose mode</span></span>

<span data-ttu-id="6cf1a-616">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-616">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="6cf1a-617">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-617">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-618">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-618">Type:</span></span>

*   <span data-ttu-id="6cf1a-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-620">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-620">Requirements</span></span>

|<span data-ttu-id="6cf1a-621">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-621">Requirement</span></span>|<span data-ttu-id="6cf1a-622">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-623">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-623">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-624">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-624">1.0</span></span>|
|[<span data-ttu-id="6cf1a-625">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-625">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-626">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-626">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-627">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-627">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-628">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-628">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-629">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-629">Example</span></span>

<span data-ttu-id="6cf1a-630">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-630">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="6cf1a-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="6cf1a-632">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-632">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="6cf1a-633">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-633">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cf1a-634">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-634">Read mode</span></span>

<span data-ttu-id="6cf1a-p131">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="6cf1a-637">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-637">Compose mode</span></span>

<span data-ttu-id="6cf1a-638">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-638">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="6cf1a-639">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-639">Type:</span></span>

*   <span data-ttu-id="6cf1a-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-641">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-641">Requirements</span></span>

|<span data-ttu-id="6cf1a-642">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-642">Requirement</span></span>|<span data-ttu-id="6cf1a-643">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-644">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-645">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-645">1.0</span></span>|
|[<span data-ttu-id="6cf1a-646">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-647">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-648">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-649">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-649">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="6cf1a-650">à : Array. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="6cf1a-651">Permet d’accéder aux destinataires de la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-651">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="6cf1a-652">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-652">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6cf1a-653">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-653">Read mode</span></span>

<span data-ttu-id="6cf1a-p133">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="6cf1a-656">Mode composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-656">Compose mode</span></span>

<span data-ttu-id="6cf1a-657">Le `to` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-657">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="6cf1a-658">Type :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-658">Type:</span></span>

*   <span data-ttu-id="6cf1a-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-660">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-660">Requirements</span></span>

|<span data-ttu-id="6cf1a-661">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-661">Requirement</span></span>|<span data-ttu-id="6cf1a-662">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-662">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-663">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-663">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-664">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-664">1.0</span></span>|
|[<span data-ttu-id="6cf1a-665">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-665">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-666">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-666">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-667">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-667">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-668">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-668">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-669">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-669">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="6cf1a-670">Méthodes</span><span class="sxs-lookup"><span data-stu-id="6cf1a-670">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="6cf1a-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6cf1a-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="6cf1a-672">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-672">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="6cf1a-673">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-673">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="6cf1a-674">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-674">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-675">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-675">Parameters:</span></span>
|<span data-ttu-id="6cf1a-676">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-676">Name</span></span>|<span data-ttu-id="6cf1a-677">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-677">Type</span></span>|<span data-ttu-id="6cf1a-678">Attributs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-678">Attributes</span></span>|<span data-ttu-id="6cf1a-679">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-679">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="6cf1a-680">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-680">String</span></span>||<span data-ttu-id="6cf1a-p134">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="6cf1a-683">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-683">String</span></span>||<span data-ttu-id="6cf1a-p135">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="6cf1a-686">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-686">Object</span></span>|<span data-ttu-id="6cf1a-687">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-687">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-688">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-688">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="6cf1a-689">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-689">Object</span></span>|<span data-ttu-id="6cf1a-690">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-690">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-691">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-691">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="6cf1a-692">Boolean</span><span class="sxs-lookup"><span data-stu-id="6cf1a-692">Boolean</span></span>|<span data-ttu-id="6cf1a-693">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-693">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-694">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-694">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="6cf1a-695">fonction</span><span class="sxs-lookup"><span data-stu-id="6cf1a-695">function</span></span>|<span data-ttu-id="6cf1a-696">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-696">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-697">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6cf1a-698">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-698">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="6cf1a-699">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-699">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6cf1a-700">Erreurs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-700">Errors</span></span>

|<span data-ttu-id="6cf1a-701">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-701">Error code</span></span>|<span data-ttu-id="6cf1a-702">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-702">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="6cf1a-703">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-703">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="6cf1a-704">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-704">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="6cf1a-705">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-705">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-706">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-706">Requirements</span></span>

|<span data-ttu-id="6cf1a-707">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-707">Requirement</span></span>|<span data-ttu-id="6cf1a-708">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-708">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-709">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-709">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-710">1.1</span><span class="sxs-lookup"><span data-stu-id="6cf1a-710">1.1</span></span>|
|[<span data-ttu-id="6cf1a-711">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-711">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-712">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-712">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cf1a-713">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-713">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-714">Composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-714">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="6cf1a-715">範例</span><span class="sxs-lookup"><span data-stu-id="6cf1a-715">Examples</span></span>

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

<span data-ttu-id="6cf1a-716">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-716">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="6cf1a-717">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6cf1a-717">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="6cf1a-718">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-718">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="6cf1a-719">Actuellement les types d’événements pris en charge sont `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, et`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="6cf1a-719">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-720">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-720">Parameters:</span></span>

| <span data-ttu-id="6cf1a-721">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-721">Name</span></span> | <span data-ttu-id="6cf1a-722">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-722">Type</span></span> | <span data-ttu-id="6cf1a-723">Attributs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-723">Attributes</span></span> | <span data-ttu-id="6cf1a-724">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-724">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="6cf1a-725">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="6cf1a-725">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="6cf1a-726">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-726">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="6cf1a-727">Fonction</span><span class="sxs-lookup"><span data-stu-id="6cf1a-727">Function</span></span> || <span data-ttu-id="6cf1a-p136">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="6cf1a-731">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-731">Object</span></span> | <span data-ttu-id="6cf1a-732">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-732">&lt;optional&gt;</span></span> | <span data-ttu-id="6cf1a-733">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-733">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="6cf1a-734">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-734">Object</span></span> | <span data-ttu-id="6cf1a-735">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-735">&lt;optional&gt;</span></span> | <span data-ttu-id="6cf1a-736">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-736">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="6cf1a-737">fonction</span><span class="sxs-lookup"><span data-stu-id="6cf1a-737">function</span></span>| <span data-ttu-id="6cf1a-738">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-738">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-739">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-739">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-740">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-740">Requirements</span></span>

|<span data-ttu-id="6cf1a-741">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-741">Requirement</span></span>| <span data-ttu-id="6cf1a-742">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-742">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-743">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-743">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cf1a-744">1.7</span><span class="sxs-lookup"><span data-stu-id="6cf1a-744">1.7</span></span> |
|[<span data-ttu-id="6cf1a-745">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-745">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cf1a-746">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-746">ReadItem</span></span> |
|[<span data-ttu-id="6cf1a-747">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-747">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cf1a-748">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-748">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="6cf1a-749">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-749">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="6cf1a-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6cf1a-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="6cf1a-751">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-751">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="6cf1a-p137">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="6cf1a-755">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-755">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="6cf1a-756">Si votre complément Office est en cours d’exécution dans Outlook Web App, le `addItemAttachmentAsync` méthode permettre joindre des éléments à des éléments autres que l’élément à modifier ; Toutefois, cela n’est pas pris en charge et n’est pas recommandée.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-756">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-757">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-757">Parameters:</span></span>

|<span data-ttu-id="6cf1a-758">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-758">Name</span></span>|<span data-ttu-id="6cf1a-759">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-759">Type</span></span>|<span data-ttu-id="6cf1a-760">Attributs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-760">Attributes</span></span>|<span data-ttu-id="6cf1a-761">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-761">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="6cf1a-762">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-762">String</span></span>||<span data-ttu-id="6cf1a-p138">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="6cf1a-765">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-765">String</span></span>||<span data-ttu-id="6cf1a-p139">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="6cf1a-768">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-768">Object</span></span>|<span data-ttu-id="6cf1a-769">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-769">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-770">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-770">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="6cf1a-771">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-771">Object</span></span>|<span data-ttu-id="6cf1a-772">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-772">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-773">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-773">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="6cf1a-774">fonction</span><span class="sxs-lookup"><span data-stu-id="6cf1a-774">function</span></span>|<span data-ttu-id="6cf1a-775">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-775">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-776">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-776">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6cf1a-777">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-777">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="6cf1a-778">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-778">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6cf1a-779">Erreurs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-779">Errors</span></span>

|<span data-ttu-id="6cf1a-780">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-780">Error code</span></span>|<span data-ttu-id="6cf1a-781">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-781">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="6cf1a-782">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-782">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-783">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-783">Requirements</span></span>

|<span data-ttu-id="6cf1a-784">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-784">Requirement</span></span>|<span data-ttu-id="6cf1a-785">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-786">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-786">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-787">1.1</span><span class="sxs-lookup"><span data-stu-id="6cf1a-787">1.1</span></span>|
|[<span data-ttu-id="6cf1a-788">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-789">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-789">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cf1a-790">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-791">Composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-791">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-792">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-792">Example</span></span>

<span data-ttu-id="6cf1a-793">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-793">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="6cf1a-794">close()</span><span class="sxs-lookup"><span data-stu-id="6cf1a-794">close()</span></span>

<span data-ttu-id="6cf1a-795">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-795">Closes the current item that is being composed.</span></span>

<span data-ttu-id="6cf1a-p140">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-798">Dans Outlook sur le web, si l’élément est un rendez-vous et il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, supprimer ou annuler même si aucune modification de l’élément depuis le dernier enregistrée.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-798">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="6cf1a-799">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-799">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-800">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-800">Requirements</span></span>

|<span data-ttu-id="6cf1a-801">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-801">Requirement</span></span>|<span data-ttu-id="6cf1a-802">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-802">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-803">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-803">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-804">1.3</span><span class="sxs-lookup"><span data-stu-id="6cf1a-804">1.3</span></span>|
|[<span data-ttu-id="6cf1a-805">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-805">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-806">Restricted</span><span class="sxs-lookup"><span data-stu-id="6cf1a-806">Restricted</span></span>|
|[<span data-ttu-id="6cf1a-807">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-807">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-808">Composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-808">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="6cf1a-809">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-809">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="6cf1a-810">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-810">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-811">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-811">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6cf1a-812">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-812">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="6cf1a-813">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-813">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="6cf1a-p141">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-817">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-817">Parameters:</span></span>

|<span data-ttu-id="6cf1a-818">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-818">Name</span></span>|<span data-ttu-id="6cf1a-819">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-819">Type</span></span>|<span data-ttu-id="6cf1a-820">Attributs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-820">Attributes</span></span>|<span data-ttu-id="6cf1a-821">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-821">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="6cf1a-822">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="6cf1a-822">String &#124; Object</span></span>||<span data-ttu-id="6cf1a-p142">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="6cf1a-825">**OU**</span><span class="sxs-lookup"><span data-stu-id="6cf1a-825">**OR**</span></span><br/><span data-ttu-id="6cf1a-p143">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="6cf1a-828">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-828">String</span></span>|<span data-ttu-id="6cf1a-829">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-829">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="6cf1a-832">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-832">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="6cf1a-833">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-833">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-834">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-834">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="6cf1a-835">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-835">String</span></span>||<span data-ttu-id="6cf1a-p145">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="6cf1a-838">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-838">String</span></span>||<span data-ttu-id="6cf1a-839">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-839">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="6cf1a-840">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-840">String</span></span>||<span data-ttu-id="6cf1a-p146">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="6cf1a-843">Boolean</span><span class="sxs-lookup"><span data-stu-id="6cf1a-843">Boolean</span></span>||<span data-ttu-id="6cf1a-p147">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="6cf1a-846">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-846">String</span></span>||<span data-ttu-id="6cf1a-p148">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="6cf1a-850">function</span><span class="sxs-lookup"><span data-stu-id="6cf1a-850">function</span></span>|<span data-ttu-id="6cf1a-851">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-851">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-852">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-852">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-853">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-853">Requirements</span></span>

|<span data-ttu-id="6cf1a-854">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-854">Requirement</span></span>|<span data-ttu-id="6cf1a-855">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-856">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-856">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-857">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-857">1.0</span></span>|
|[<span data-ttu-id="6cf1a-858">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-859">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-859">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-860">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-861">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-861">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="6cf1a-862">Exemples</span><span class="sxs-lookup"><span data-stu-id="6cf1a-862">Examples</span></span>

<span data-ttu-id="6cf1a-863">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-863">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="6cf1a-864">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-864">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="6cf1a-865">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-865">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="6cf1a-866">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-866">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="6cf1a-867">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-867">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="6cf1a-868">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-868">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="6cf1a-869">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-869">displayReplyForm(formData)</span></span>

<span data-ttu-id="6cf1a-870">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-870">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-871">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-871">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6cf1a-872">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-872">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="6cf1a-873">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-873">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="6cf1a-p149">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-877">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-877">Parameters:</span></span>

|<span data-ttu-id="6cf1a-878">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-878">Name</span></span>|<span data-ttu-id="6cf1a-879">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-879">Type</span></span>|<span data-ttu-id="6cf1a-880">Attributs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-880">Attributes</span></span>|<span data-ttu-id="6cf1a-881">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-881">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="6cf1a-882">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="6cf1a-882">String &#124; Object</span></span>||<span data-ttu-id="6cf1a-p150">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="6cf1a-885">**OU**</span><span class="sxs-lookup"><span data-stu-id="6cf1a-885">**OR**</span></span><br/><span data-ttu-id="6cf1a-p151">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="6cf1a-888">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-888">String</span></span>|<span data-ttu-id="6cf1a-889">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-889">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-p152">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="6cf1a-892">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-892">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="6cf1a-893">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-893">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-894">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-894">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="6cf1a-895">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-895">String</span></span>||<span data-ttu-id="6cf1a-p153">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="6cf1a-898">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-898">String</span></span>||<span data-ttu-id="6cf1a-899">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-899">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="6cf1a-900">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-900">String</span></span>||<span data-ttu-id="6cf1a-p154">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="6cf1a-903">Boolean</span><span class="sxs-lookup"><span data-stu-id="6cf1a-903">Boolean</span></span>||<span data-ttu-id="6cf1a-p155">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="6cf1a-906">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-906">String</span></span>||<span data-ttu-id="6cf1a-p156">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="6cf1a-910">function</span><span class="sxs-lookup"><span data-stu-id="6cf1a-910">function</span></span>|<span data-ttu-id="6cf1a-911">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-911">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-912">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-912">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-913">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-913">Requirements</span></span>

|<span data-ttu-id="6cf1a-914">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-914">Requirement</span></span>|<span data-ttu-id="6cf1a-915">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-915">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-916">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-916">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-917">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-917">1.0</span></span>|
|[<span data-ttu-id="6cf1a-918">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-918">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-919">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-919">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-920">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-920">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-921">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-921">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="6cf1a-922">Exemples</span><span class="sxs-lookup"><span data-stu-id="6cf1a-922">Examples</span></span>

<span data-ttu-id="6cf1a-923">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-923">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="6cf1a-924">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-924">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="6cf1a-925">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-925">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="6cf1a-926">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-926">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="6cf1a-927">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-927">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="6cf1a-928">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-928">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="6cf1a-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="6cf1a-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="6cf1a-930">Obtient les entités trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-930">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-931">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-931">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-932">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-932">Requirements</span></span>

|<span data-ttu-id="6cf1a-933">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-933">Requirement</span></span>|<span data-ttu-id="6cf1a-934">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-935">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-936">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-936">1.0</span></span>|
|[<span data-ttu-id="6cf1a-937">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-938">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-939">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-940">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cf1a-941">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-941">Returns:</span></span>

<span data-ttu-id="6cf1a-942">Type : [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-942">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="6cf1a-943">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-943">Example</span></span>

<span data-ttu-id="6cf1a-944">L’exemple suivant accède aux entités de contacts dans le corps de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-944">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="6cf1a-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="6cf1a-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="6cf1a-946">Obtient un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-946">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-947">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-947">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-948">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-948">Parameters:</span></span>

|<span data-ttu-id="6cf1a-949">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-949">Name</span></span>|<span data-ttu-id="6cf1a-950">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-950">Type</span></span>|<span data-ttu-id="6cf1a-951">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-951">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="6cf1a-952">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="6cf1a-952">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="6cf1a-953">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-953">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-954">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-954">Requirements</span></span>

|<span data-ttu-id="6cf1a-955">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-955">Requirement</span></span>|<span data-ttu-id="6cf1a-956">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-957">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-957">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-958">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-958">1.0</span></span>|
|[<span data-ttu-id="6cf1a-959">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-960">Restricted</span><span class="sxs-lookup"><span data-stu-id="6cf1a-960">Restricted</span></span>|
|[<span data-ttu-id="6cf1a-961">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-962">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cf1a-963">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-963">Returns:</span></span>

<span data-ttu-id="6cf1a-964">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-964">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="6cf1a-965">Si aucune entité du type spécifié n’est présentes dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-965">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="6cf1a-966">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-966">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="6cf1a-967">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-967">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="6cf1a-968">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="6cf1a-968">Value of `entityType`</span></span>|<span data-ttu-id="6cf1a-969">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="6cf1a-969">Type of objects in returned array</span></span>|<span data-ttu-id="6cf1a-970">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="6cf1a-970">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="6cf1a-971">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-971">String</span></span>|<span data-ttu-id="6cf1a-972">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="6cf1a-972">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="6cf1a-973">Contact</span><span class="sxs-lookup"><span data-stu-id="6cf1a-973">Contact</span></span>|<span data-ttu-id="6cf1a-974">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6cf1a-974">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="6cf1a-975">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-975">String</span></span>|<span data-ttu-id="6cf1a-976">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6cf1a-976">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="6cf1a-977">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="6cf1a-977">MeetingSuggestion</span></span>|<span data-ttu-id="6cf1a-978">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6cf1a-978">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="6cf1a-979">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="6cf1a-979">PhoneNumber</span></span>|<span data-ttu-id="6cf1a-980">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="6cf1a-980">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="6cf1a-981">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="6cf1a-981">TaskSuggestion</span></span>|<span data-ttu-id="6cf1a-982">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6cf1a-982">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="6cf1a-983">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-983">String</span></span>|<span data-ttu-id="6cf1a-984">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="6cf1a-984">**Restricted**</span></span>|

<span data-ttu-id="6cf1a-985">Type : Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="6cf1a-985">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="6cf1a-986">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-986">Example</span></span>

<span data-ttu-id="6cf1a-987">L’exemple suivant montre comment accéder à un tableau de chaînes qui représentent les adresses postales dans le corps de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-987">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="6cf1a-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="6cf1a-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="6cf1a-989">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-989">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-990">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-990">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6cf1a-991">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-991">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-992">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-992">Parameters:</span></span>

|<span data-ttu-id="6cf1a-993">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-993">Name</span></span>|<span data-ttu-id="6cf1a-994">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-994">Type</span></span>|<span data-ttu-id="6cf1a-995">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-995">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="6cf1a-996">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-996">String</span></span>|<span data-ttu-id="6cf1a-997">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-997">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-998">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-998">Requirements</span></span>

|<span data-ttu-id="6cf1a-999">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-999">Requirement</span></span>|<span data-ttu-id="6cf1a-1000">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1000">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-1001">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1001">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-1002">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1002">1.0</span></span>|
|[<span data-ttu-id="6cf1a-1003">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1003">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-1004">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1004">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-1005">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1005">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-1006">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1006">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cf1a-1007">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1007">Returns:</span></span>

<span data-ttu-id="6cf1a-p158">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="6cf1a-1010">Type : Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="6cf1a-1010">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="6cf1a-1011">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1011">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="6cf1a-1012">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1012">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-1013">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1013">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6cf1a-p159">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="6cf1a-1017">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1017">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="6cf1a-1018">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1018">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="6cf1a-p160">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-1022">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1022">Requirements</span></span>

|<span data-ttu-id="6cf1a-1023">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1023">Requirement</span></span>|<span data-ttu-id="6cf1a-1024">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1024">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-1025">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1025">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-1026">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1026">1.0</span></span>|
|[<span data-ttu-id="6cf1a-1027">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1027">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-1028">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1028">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-1029">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1029">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-1030">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1030">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cf1a-1031">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1031">Returns:</span></span>

<span data-ttu-id="6cf1a-p161">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="6cf1a-1034">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1034">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6cf1a-1035">Object</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1035">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6cf1a-1036">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1036">Example</span></span>

<span data-ttu-id="6cf1a-1037">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1037">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="6cf1a-1038">getregexmatchesbyname (Name) → (nullable) {tableau. < chaîne >}</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="6cf1a-1039">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1039">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-1040">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1040">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6cf1a-1041">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1041">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="6cf1a-p162">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-1044">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1044">Parameters:</span></span>

|<span data-ttu-id="6cf1a-1045">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1045">Name</span></span>|<span data-ttu-id="6cf1a-1046">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1046">Type</span></span>|<span data-ttu-id="6cf1a-1047">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1047">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="6cf1a-1048">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1048">String</span></span>|<span data-ttu-id="6cf1a-1049">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1049">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-1050">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1050">Requirements</span></span>

|<span data-ttu-id="6cf1a-1051">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1051">Requirement</span></span>|<span data-ttu-id="6cf1a-1052">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-1053">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1053">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-1054">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1054">1.0</span></span>|
|[<span data-ttu-id="6cf1a-1055">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1056">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-1057">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-1058">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1058">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cf1a-1059">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1059">Returns:</span></span>

<span data-ttu-id="6cf1a-1060">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1060">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="6cf1a-1061">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1061">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6cf1a-1062">< Chaîne > tableau.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1062">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6cf1a-1063">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1063">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="6cf1a-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="6cf1a-1065">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1065">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="6cf1a-p163">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-1068">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1068">Parameters:</span></span>

|<span data-ttu-id="6cf1a-1069">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1069">Name</span></span>|<span data-ttu-id="6cf1a-1070">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1070">Type</span></span>|<span data-ttu-id="6cf1a-1071">Attributs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1071">Attributes</span></span>|<span data-ttu-id="6cf1a-1072">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1072">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="6cf1a-1073">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1073">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="6cf1a-p164">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="6cf1a-1077">Object</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1077">Object</span></span>|<span data-ttu-id="6cf1a-1078">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1078">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-1079">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1079">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="6cf1a-1080">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1080">Object</span></span>|<span data-ttu-id="6cf1a-1081">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1081">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-1082">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1082">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="6cf1a-1083">fonction</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1083">function</span></span>||<span data-ttu-id="6cf1a-1084">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1084">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6cf1a-1085">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1085">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="6cf1a-1086">Pour accéder à la propriété source provenant de la sélection, appelez `asyncResult.value.sourceProperty`, qui sera soit `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1086">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-1087">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1087">Requirements</span></span>

|<span data-ttu-id="6cf1a-1088">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1088">Requirement</span></span>|<span data-ttu-id="6cf1a-1089">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-1090">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1090">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-1091">1.2</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1091">1.2</span></span>|
|[<span data-ttu-id="6cf1a-1092">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-1093">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1093">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cf1a-1094">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-1095">Composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1095">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cf1a-1096">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1096">Returns:</span></span>

<span data-ttu-id="6cf1a-1097">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1097">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="6cf1a-1098">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1098">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6cf1a-1099">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1099">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6cf1a-1100">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1100">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="6cf1a-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="6cf1a-p166">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-1104">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1104">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-1105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1105">Requirements</span></span>

|<span data-ttu-id="6cf1a-1106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1106">Requirement</span></span>|<span data-ttu-id="6cf1a-1107">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1107">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-1108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-1109">1.6</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1109">1.6</span></span>|
|[<span data-ttu-id="6cf1a-1110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-1111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1111">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-1112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-1113">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1113">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cf1a-1114">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1114">Returns:</span></span>

<span data-ttu-id="6cf1a-1115">Type : [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1115">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="6cf1a-1116">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1116">Example</span></span>

<span data-ttu-id="6cf1a-1117">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1117">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="6cf1a-1118">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1118">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="6cf1a-p167">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-1121">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1121">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6cf1a-p168">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="6cf1a-1125">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1125">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="6cf1a-1126">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1126">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="6cf1a-p169">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cf1a-1130">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1130">Requirements</span></span>

|<span data-ttu-id="6cf1a-1131">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1131">Requirement</span></span>|<span data-ttu-id="6cf1a-1132">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-1133">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1133">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-1134">1.6</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1134">1.6</span></span>|
|[<span data-ttu-id="6cf1a-1135">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1136">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-1137">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-1138">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1138">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6cf1a-1139">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1139">Returns:</span></span>

<span data-ttu-id="6cf1a-p170">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="6cf1a-1142">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1142">Example</span></span>

<span data-ttu-id="6cf1a-1143">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1143">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="6cf1a-1144">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1144">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="6cf1a-1145">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1145">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="6cf1a-p171">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-1149">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1149">Parameters:</span></span>

|<span data-ttu-id="6cf1a-1150">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1150">Name</span></span>|<span data-ttu-id="6cf1a-1151">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1151">Type</span></span>|<span data-ttu-id="6cf1a-1152">Attributs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1152">Attributes</span></span>|<span data-ttu-id="6cf1a-1153">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1153">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="6cf1a-1154">function</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1154">function</span></span>||<span data-ttu-id="6cf1a-1155">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1155">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6cf1a-1156">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1156">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="6cf1a-1157">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées de l’élément et enregistrer les modifications à la propriété personnalisée est redéfinie sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1157">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="6cf1a-1158">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1158">Object</span></span>|<span data-ttu-id="6cf1a-1159">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-1160">Les développeurs peuvent fournir n’importe quel objet qu’ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1160">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="6cf1a-1161">Cet objet est accessible par le `asyncResult.asyncContext` propriété dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1161">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-1162">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1162">Requirements</span></span>

|<span data-ttu-id="6cf1a-1163">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1163">Requirement</span></span>|<span data-ttu-id="6cf1a-1164">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1164">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-1165">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1165">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-1166">1.0</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1166">1.0</span></span>|
|[<span data-ttu-id="6cf1a-1167">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1167">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-1168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1168">ReadItem</span></span>|
|[<span data-ttu-id="6cf1a-1169">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1169">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-1170">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1170">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-1171">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1171">Example</span></span>

<span data-ttu-id="6cf1a-p174">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="6cf1a-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="6cf1a-1176">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1176">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="6cf1a-p175">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-1181">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1181">Parameters:</span></span>

|<span data-ttu-id="6cf1a-1182">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1182">Name</span></span>|<span data-ttu-id="6cf1a-1183">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1183">Type</span></span>|<span data-ttu-id="6cf1a-1184">Attributs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1184">Attributes</span></span>|<span data-ttu-id="6cf1a-1185">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1185">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="6cf1a-1186">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1186">String</span></span>||<span data-ttu-id="6cf1a-p176">Identificateur de la pièce jointe à supprimer. La longueur maximale de la chaîne est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p176">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="6cf1a-1189">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1189">Object</span></span>|<span data-ttu-id="6cf1a-1190">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-1191">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="6cf1a-1192">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1192">Object</span></span>|<span data-ttu-id="6cf1a-1193">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-1194">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="6cf1a-1195">fonction</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1195">function</span></span>|<span data-ttu-id="6cf1a-1196">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-1197">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6cf1a-1198">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6cf1a-1199">Erreurs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1199">Errors</span></span>

|<span data-ttu-id="6cf1a-1200">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1200">Error code</span></span>|<span data-ttu-id="6cf1a-1201">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="6cf1a-1202">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-1203">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1203">Requirements</span></span>

|<span data-ttu-id="6cf1a-1204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1204">Requirement</span></span>|<span data-ttu-id="6cf1a-1205">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-1206">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1206">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1207">1.1</span></span>|
|[<span data-ttu-id="6cf1a-1208">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cf1a-1210">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-1211">Composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-1212">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1212">Example</span></span>

<span data-ttu-id="6cf1a-1213">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1213">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="6cf1a-1214">removeHandlerAsync (eventType, gestionnaire, [options], [rappel])</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1214">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="6cf1a-1215">Supprime un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1215">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="6cf1a-1216">Actuellement les types d’événements pris en charge sont `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, et`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-1217">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1217">Parameters:</span></span>

| <span data-ttu-id="6cf1a-1218">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1218">Name</span></span> | <span data-ttu-id="6cf1a-1219">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1219">Type</span></span> | <span data-ttu-id="6cf1a-1220">Attributs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1220">Attributes</span></span> | <span data-ttu-id="6cf1a-1221">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="6cf1a-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="6cf1a-1223">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1223">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="6cf1a-1224">Fonction</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1224">Function</span></span> || <span data-ttu-id="6cf1a-p177">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p177">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="6cf1a-1228">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1228">Object</span></span> | <span data-ttu-id="6cf1a-1229">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="6cf1a-1230">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="6cf1a-1231">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1231">Object</span></span> | <span data-ttu-id="6cf1a-1232">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="6cf1a-1233">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="6cf1a-1234">fonction</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1234">function</span></span>| <span data-ttu-id="6cf1a-1235">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-1236">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-1237">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1237">Requirements</span></span>

|<span data-ttu-id="6cf1a-1238">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1238">Requirement</span></span>| <span data-ttu-id="6cf1a-1239">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-1240">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6cf1a-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1241">1.7</span></span> |
|[<span data-ttu-id="6cf1a-1242">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6cf1a-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1243">ReadItem</span></span> |
|[<span data-ttu-id="6cf1a-1244">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6cf1a-1245">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1245">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="6cf1a-1246">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1246">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="6cf1a-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="6cf1a-1248">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="6cf1a-p178">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-1252">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` pour utiliser avec EWS ou l’API REST, gardez à l’esprit que quand Outlook est en mode mis en cache, il peut prendre un certain temps avant que l’élément est réellement synchronisé avec le serveur.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="6cf1a-1253">Jusqu'à ce que l’élément est synchronisé, à l’aide de la `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="6cf1a-p180">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="6cf1a-1257">Les clients suivants ont un comportement différent pour `saveAsync` de rendez-vous dans le mode de composition :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="6cf1a-1258">Mac Outlook ne gère pas `saveAsync` d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1258">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="6cf1a-1259">Appel de `saveAsync` d’une réunion dans Outlook Mac renverra une erreur.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1259">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="6cf1a-1260">Outlook sur le web toujours envoie une invitation ou mettre à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-1261">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1261">Parameters:</span></span>

|<span data-ttu-id="6cf1a-1262">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1262">Name</span></span>|<span data-ttu-id="6cf1a-1263">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1263">Type</span></span>|<span data-ttu-id="6cf1a-1264">Attributs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1264">Attributes</span></span>|<span data-ttu-id="6cf1a-1265">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="6cf1a-1266">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1266">Object</span></span>|<span data-ttu-id="6cf1a-1267">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-1268">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="6cf1a-1269">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1269">Object</span></span>|<span data-ttu-id="6cf1a-1270">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-1271">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="6cf1a-1272">fonction</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1272">function</span></span>||<span data-ttu-id="6cf1a-1273">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6cf1a-1274">En cas de réussite, l’identificateur d’élément est fournie dans le `asyncResult.value` propriété.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1274">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-1275">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1275">Requirements</span></span>

|<span data-ttu-id="6cf1a-1276">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1276">Requirement</span></span>|<span data-ttu-id="6cf1a-1277">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-1278">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1279">1.3</span></span>|
|[<span data-ttu-id="6cf1a-1280">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cf1a-1282">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-1283">Composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="6cf1a-1284">範例</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1284">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="6cf1a-p182">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="6cf1a-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="6cf1a-1288">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="6cf1a-p183">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6cf1a-1292">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1292">Parameters:</span></span>

|<span data-ttu-id="6cf1a-1293">Nom</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1293">Name</span></span>|<span data-ttu-id="6cf1a-1294">Type</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1294">Type</span></span>|<span data-ttu-id="6cf1a-1295">Attributs</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1295">Attributes</span></span>|<span data-ttu-id="6cf1a-1296">Description</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="6cf1a-1297">String</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1297">String</span></span>||<span data-ttu-id="6cf1a-p184">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="6cf1a-1301">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1301">Object</span></span>|<span data-ttu-id="6cf1a-1302">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-1303">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="6cf1a-1304">Objet</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1304">Object</span></span>|<span data-ttu-id="6cf1a-1305">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-1306">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="6cf1a-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="6cf1a-1308">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="6cf1a-p185">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="6cf1a-p186">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="6cf1a-1313">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="6cf1a-1314">fonction</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1314">function</span></span>||<span data-ttu-id="6cf1a-1315">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6cf1a-1316">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1316">Requirements</span></span>

|<span data-ttu-id="6cf1a-1317">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1317">Requirement</span></span>|<span data-ttu-id="6cf1a-1318">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cf1a-1319">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1319">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="6cf1a-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1320">1.2</span></span>|
|[<span data-ttu-id="6cf1a-1321">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="6cf1a-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="6cf1a-1323">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="6cf1a-1324">Composition</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6cf1a-1325">Exemples</span><span class="sxs-lookup"><span data-stu-id="6cf1a-1325">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```