
# <a name="item"></a><span data-ttu-id="8e1c2-101">élément</span><span class="sxs-lookup"><span data-stu-id="8e1c2-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="8e1c2-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="8e1c2-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="8e1c2-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-105">Requirements</span></span>

|<span data-ttu-id="8e1c2-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-106">Requirement</span></span>|<span data-ttu-id="8e1c2-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-109">1.0</span></span>|
|[<span data-ttu-id="8e1c2-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-111">Restricted</span><span class="sxs-lookup"><span data-stu-id="8e1c2-111">Restricted</span></span>|
|[<span data-ttu-id="8e1c2-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8e1c2-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="8e1c2-114">Members and methods</span></span>

| <span data-ttu-id="8e1c2-115">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-115">Member</span></span> | <span data-ttu-id="8e1c2-116">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8e1c2-117">attachments</span><span class="sxs-lookup"><span data-stu-id="8e1c2-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="8e1c2-118">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-118">Member</span></span> |
| [<span data-ttu-id="8e1c2-119">bcc</span><span class="sxs-lookup"><span data-stu-id="8e1c2-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8e1c2-120">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-120">Member</span></span> |
| [<span data-ttu-id="8e1c2-121">body</span><span class="sxs-lookup"><span data-stu-id="8e1c2-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="8e1c2-122">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-122">Member</span></span> |
| [<span data-ttu-id="8e1c2-123">cc</span><span class="sxs-lookup"><span data-stu-id="8e1c2-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8e1c2-124">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-124">Member</span></span> |
| [<span data-ttu-id="8e1c2-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="8e1c2-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="8e1c2-126">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-126">Member</span></span> |
| [<span data-ttu-id="8e1c2-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="8e1c2-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="8e1c2-128">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-128">Member</span></span> |
| [<span data-ttu-id="8e1c2-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="8e1c2-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="8e1c2-130">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-130">Member</span></span> |
| [<span data-ttu-id="8e1c2-131">end</span><span class="sxs-lookup"><span data-stu-id="8e1c2-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="8e1c2-132">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-132">Member</span></span> |
| [<span data-ttu-id="8e1c2-133">from</span><span class="sxs-lookup"><span data-stu-id="8e1c2-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="8e1c2-134">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-134">Member</span></span> |
| [<span data-ttu-id="8e1c2-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="8e1c2-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="8e1c2-136">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-136">Member</span></span> |
| [<span data-ttu-id="8e1c2-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="8e1c2-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="8e1c2-138">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-138">Member</span></span> |
| [<span data-ttu-id="8e1c2-139">itemId</span><span class="sxs-lookup"><span data-stu-id="8e1c2-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="8e1c2-140">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-140">Member</span></span> |
| [<span data-ttu-id="8e1c2-141">itemType</span><span class="sxs-lookup"><span data-stu-id="8e1c2-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="8e1c2-142">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-142">Member</span></span> |
| [<span data-ttu-id="8e1c2-143">location</span><span class="sxs-lookup"><span data-stu-id="8e1c2-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="8e1c2-144">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-144">Member</span></span> |
| [<span data-ttu-id="8e1c2-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="8e1c2-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="8e1c2-146">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-146">Member</span></span> |
| [<span data-ttu-id="8e1c2-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="8e1c2-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="8e1c2-148">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-148">Member</span></span> |
| [<span data-ttu-id="8e1c2-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="8e1c2-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8e1c2-150">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-150">Member</span></span> |
| [<span data-ttu-id="8e1c2-151">organizer</span><span class="sxs-lookup"><span data-stu-id="8e1c2-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="8e1c2-152">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-152">Member</span></span> |
| [<span data-ttu-id="8e1c2-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="8e1c2-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="8e1c2-154">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-154">Member</span></span> |
| [<span data-ttu-id="8e1c2-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="8e1c2-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8e1c2-156">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-156">Member</span></span> |
| [<span data-ttu-id="8e1c2-157">sender</span><span class="sxs-lookup"><span data-stu-id="8e1c2-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="8e1c2-158">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-158">Member</span></span> |
| [<span data-ttu-id="8e1c2-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="8e1c2-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="8e1c2-160">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-160">Member</span></span> |
| [<span data-ttu-id="8e1c2-161">start</span><span class="sxs-lookup"><span data-stu-id="8e1c2-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="8e1c2-162">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-162">Member</span></span> |
| [<span data-ttu-id="8e1c2-163">subject</span><span class="sxs-lookup"><span data-stu-id="8e1c2-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="8e1c2-164">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-164">Member</span></span> |
| [<span data-ttu-id="8e1c2-165">to</span><span class="sxs-lookup"><span data-stu-id="8e1c2-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8e1c2-166">Membre</span><span class="sxs-lookup"><span data-stu-id="8e1c2-166">Member</span></span> |
| [<span data-ttu-id="8e1c2-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8e1c2-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="8e1c2-168">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-168">Method</span></span> |
| [<span data-ttu-id="8e1c2-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="8e1c2-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="8e1c2-170">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-170">Method</span></span> |
| [<span data-ttu-id="8e1c2-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="8e1c2-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="8e1c2-172">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-172">Method</span></span> |
| [<span data-ttu-id="8e1c2-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8e1c2-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="8e1c2-174">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-174">Method</span></span> |
| [<span data-ttu-id="8e1c2-175">close</span><span class="sxs-lookup"><span data-stu-id="8e1c2-175">close</span></span>](#close) | <span data-ttu-id="8e1c2-176">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-176">Method</span></span> |
| [<span data-ttu-id="8e1c2-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="8e1c2-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="8e1c2-178">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-178">Method</span></span> |
| [<span data-ttu-id="8e1c2-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="8e1c2-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="8e1c2-180">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-180">Method</span></span> |
| [<span data-ttu-id="8e1c2-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="8e1c2-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="8e1c2-182">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-182">Method</span></span> |
| [<span data-ttu-id="8e1c2-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="8e1c2-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="8e1c2-184">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-184">Method</span></span> |
| [<span data-ttu-id="8e1c2-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="8e1c2-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="8e1c2-186">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-186">Method</span></span> |
| [<span data-ttu-id="8e1c2-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="8e1c2-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="8e1c2-188">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-188">Method</span></span> |
| [<span data-ttu-id="8e1c2-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8e1c2-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="8e1c2-190">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-190">Method</span></span> |
| [<span data-ttu-id="8e1c2-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="8e1c2-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="8e1c2-192">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-192">Method</span></span> |
| [<span data-ttu-id="8e1c2-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8e1c2-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="8e1c2-194">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-194">Method</span></span> |
| [<span data-ttu-id="8e1c2-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="8e1c2-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="8e1c2-196">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-196">Method</span></span> |
| [<span data-ttu-id="8e1c2-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8e1c2-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="8e1c2-198">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-198">Method</span></span> |
| [<span data-ttu-id="8e1c2-199">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8e1c2-199">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="8e1c2-200">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-200">Method</span></span> |
| [<span data-ttu-id="8e1c2-201">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8e1c2-201">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="8e1c2-202">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-202">Method</span></span> |
| [<span data-ttu-id="8e1c2-203">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8e1c2-203">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="8e1c2-204">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-204">Method</span></span> |
| [<span data-ttu-id="8e1c2-205">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="8e1c2-205">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="8e1c2-206">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-206">Method</span></span> |
| [<span data-ttu-id="8e1c2-207">saveAsync</span><span class="sxs-lookup"><span data-stu-id="8e1c2-207">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="8e1c2-208">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-208">Method</span></span> |
| [<span data-ttu-id="8e1c2-209">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8e1c2-209">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="8e1c2-210">Méthode</span><span class="sxs-lookup"><span data-stu-id="8e1c2-210">Method</span></span> |

### <a name="example"></a><span data-ttu-id="8e1c2-211">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-211">Example</span></span>

<span data-ttu-id="8e1c2-212">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-212">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="8e1c2-213">Membres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-213">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="8e1c2-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8e1c2-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="8e1c2-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-217">Certains types de fichiers bloqués par Outlook en raison de problèmes de sécurité potentiels et ne sont donc pas retournés.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-217">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8e1c2-218">Pour plus d’informations, voir les [pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-218">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-219">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-219">Type:</span></span>

*   <span data-ttu-id="8e1c2-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8e1c2-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-221">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-221">Requirements</span></span>

|<span data-ttu-id="8e1c2-222">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-222">Requirement</span></span>|<span data-ttu-id="8e1c2-223">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-224">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-225">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-225">1.0</span></span>|
|[<span data-ttu-id="8e1c2-226">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-227">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-228">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-229">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-230">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-230">Example</span></span>

<span data-ttu-id="8e1c2-231">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-231">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8e1c2-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8e1c2-233">Obtient un objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires des Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-233">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8e1c2-234">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-234">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-235">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-235">Type:</span></span>

*   [<span data-ttu-id="8e1c2-236">Destinataires</span><span class="sxs-lookup"><span data-stu-id="8e1c2-236">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="8e1c2-237">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-237">Requirements</span></span>

|<span data-ttu-id="8e1c2-238">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-238">Requirement</span></span>|<span data-ttu-id="8e1c2-239">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-240">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-241">1.1</span><span class="sxs-lookup"><span data-stu-id="8e1c2-241">1.1</span></span>|
|[<span data-ttu-id="8e1c2-242">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-243">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-244">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-245">Composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-245">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-246">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-246">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="8e1c2-247">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-247">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="8e1c2-248">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-248">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-249">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-249">Type:</span></span>

*   [<span data-ttu-id="8e1c2-250">Corps</span><span class="sxs-lookup"><span data-stu-id="8e1c2-250">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="8e1c2-251">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-251">Requirements</span></span>

|<span data-ttu-id="8e1c2-252">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-252">Requirement</span></span>|<span data-ttu-id="8e1c2-253">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-254">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-254">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-255">1.1</span><span class="sxs-lookup"><span data-stu-id="8e1c2-255">1.1</span></span>|
|[<span data-ttu-id="8e1c2-256">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-256">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-257">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-258">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-258">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-259">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-259">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8e1c2-260">cc : tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-260">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8e1c2-261">Permet d’accéder aux destinataires Cc (copie carbone) d’un message.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-261">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8e1c2-262">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-262">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e1c2-263">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-263">Read mode</span></span>

<span data-ttu-id="8e1c2-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e1c2-266">Mode composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-266">Compose mode</span></span>

<span data-ttu-id="8e1c2-267">Le `cc` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-267">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-268">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-268">Type:</span></span>

*   <span data-ttu-id="8e1c2-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-270">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-270">Requirements</span></span>

|<span data-ttu-id="8e1c2-271">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-271">Requirement</span></span>|<span data-ttu-id="8e1c2-272">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-273">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-274">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-274">1.0</span></span>|
|[<span data-ttu-id="8e1c2-275">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-276">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-277">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-278">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-278">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-279">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-279">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="8e1c2-280">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-280">(nullable) conversationId :String</span></span>

<span data-ttu-id="8e1c2-281">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-281">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8e1c2-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8e1c2-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-286">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-286">Type:</span></span>

*   <span data-ttu-id="8e1c2-287">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8e1c2-287">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-288">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-288">Requirements</span></span>

|<span data-ttu-id="8e1c2-289">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-289">Requirement</span></span>|<span data-ttu-id="8e1c2-290">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-291">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-292">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-292">1.0</span></span>|
|[<span data-ttu-id="8e1c2-293">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-294">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-295">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-296">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-296">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="8e1c2-297">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="8e1c2-297">dateTimeCreated :Date</span></span>

<span data-ttu-id="8e1c2-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-300">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-300">Type:</span></span>

*   <span data-ttu-id="8e1c2-301">Date</span><span class="sxs-lookup"><span data-stu-id="8e1c2-301">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-302">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-302">Requirements</span></span>

|<span data-ttu-id="8e1c2-303">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-303">Requirement</span></span>|<span data-ttu-id="8e1c2-304">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-305">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-305">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-306">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-306">1.0</span></span>|
|[<span data-ttu-id="8e1c2-307">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-308">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-309">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-310">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-311">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-311">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="8e1c2-312">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="8e1c2-312">dateTimeModified :Date</span></span>

<span data-ttu-id="8e1c2-p110">Obtient la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-315">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-315">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-316">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-316">Type:</span></span>

*   <span data-ttu-id="8e1c2-317">Date</span><span class="sxs-lookup"><span data-stu-id="8e1c2-317">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-318">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-318">Requirements</span></span>

|<span data-ttu-id="8e1c2-319">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-319">Requirement</span></span>|<span data-ttu-id="8e1c2-320">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-320">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-321">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-321">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-322">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-322">1.0</span></span>|
|[<span data-ttu-id="8e1c2-323">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-323">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-324">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-324">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-325">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-325">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-326">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-326">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-327">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-327">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="8e1c2-328">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-328">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="8e1c2-329">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-329">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8e1c2-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e1c2-332">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-332">Read mode</span></span>

<span data-ttu-id="8e1c2-333">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-333">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e1c2-334">Mode composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-334">Compose mode</span></span>

<span data-ttu-id="8e1c2-335">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-335">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8e1c2-336">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-336">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-337">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-337">Type:</span></span>

*   <span data-ttu-id="8e1c2-338">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-338">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-339">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-339">Requirements</span></span>

|<span data-ttu-id="8e1c2-340">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-340">Requirement</span></span>|<span data-ttu-id="8e1c2-341">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-342">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-342">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-343">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-343">1.0</span></span>|
|[<span data-ttu-id="8e1c2-344">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-345">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-346">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-347">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-348">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-348">Example</span></span>

<span data-ttu-id="8e1c2-349">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-349">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="8e1c2-350">à partir de :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[à partir de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-350">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="8e1c2-351">Obtient l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-351">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="8e1c2-p112">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-354">Le `recipientType` propriété de la `EmailAddressDetails` objet dans le `from` propriété est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-354">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e1c2-355">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-355">Read mode</span></span>

<span data-ttu-id="8e1c2-356">Le `from` propriété retourne un `EmailAddressDetails` objet.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-356">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="8e1c2-357">Mode composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-357">Compose mode</span></span>

<span data-ttu-id="8e1c2-358">Le `from` propriété renvoie une `From` objet qui fournit une méthode pour obtenir la valeur affectée à.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-358">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8e1c2-359">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-359">Type:</span></span>

*   <span data-ttu-id="8e1c2-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [à partir de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-361">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-361">Requirements</span></span>

|<span data-ttu-id="8e1c2-362">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-362">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="8e1c2-363">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-363">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-364">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-364">1.0</span></span>|<span data-ttu-id="8e1c2-365">1.7</span><span class="sxs-lookup"><span data-stu-id="8e1c2-365">1.7</span></span>|
|[<span data-ttu-id="8e1c2-366">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-366">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-367">ReadItem</span></span>|<span data-ttu-id="8e1c2-368">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-368">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e1c2-369">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-369">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-370">Read</span><span class="sxs-lookup"><span data-stu-id="8e1c2-370">Read</span></span>|<span data-ttu-id="8e1c2-371">Composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-371">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="8e1c2-372">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-372">internetMessageId :String</span></span>

<span data-ttu-id="8e1c2-p113">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-375">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-375">Type:</span></span>

*   <span data-ttu-id="8e1c2-376">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8e1c2-376">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-377">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-377">Requirements</span></span>

|<span data-ttu-id="8e1c2-378">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-378">Requirement</span></span>|<span data-ttu-id="8e1c2-379">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-380">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-380">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-381">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-381">1.0</span></span>|
|[<span data-ttu-id="8e1c2-382">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-382">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-383">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-384">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-384">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-385">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-385">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-386">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-386">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="8e1c2-387">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-387">itemClass :String</span></span>

<span data-ttu-id="8e1c2-p114">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8e1c2-p115">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="8e1c2-392">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-392">Type</span></span>|<span data-ttu-id="8e1c2-393">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-393">Description</span></span>|<span data-ttu-id="8e1c2-394">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="8e1c2-394">item class</span></span>|
|---|---|---|
|<span data-ttu-id="8e1c2-395">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="8e1c2-395">Appointment items</span></span>|<span data-ttu-id="8e1c2-396">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-396">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="8e1c2-397">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="8e1c2-397">Message items</span></span>|<span data-ttu-id="8e1c2-398">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-398">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="8e1c2-399">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-399">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-400">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-400">Type:</span></span>

*   <span data-ttu-id="8e1c2-401">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8e1c2-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-402">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-402">Requirements</span></span>

|<span data-ttu-id="8e1c2-403">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-403">Requirement</span></span>|<span data-ttu-id="8e1c2-404">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-405">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-405">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-406">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-406">1.0</span></span>|
|[<span data-ttu-id="8e1c2-407">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-407">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-408">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-409">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-409">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-410">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-411">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-411">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8e1c2-412">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-412">(nullable) itemId :String</span></span>

<span data-ttu-id="8e1c2-p116">Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-415">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-415">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8e1c2-416">Le `itemId` propriété n’est pas identique à l’identificateur d’entrée Outlook ou l’ID utilisé par l’API REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-416">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8e1c2-417">Avant d’effectuer des appels d’API REST à l’aide de cette valeur, elle doit être convertie à l’aide de [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-417">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8e1c2-418">Pour plus d’informations, voir [utiliser les API REST d’Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-418">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="8e1c2-p118">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-421">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-421">Type:</span></span>

*   <span data-ttu-id="8e1c2-422">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8e1c2-422">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-423">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-423">Requirements</span></span>

|<span data-ttu-id="8e1c2-424">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-424">Requirement</span></span>|<span data-ttu-id="8e1c2-425">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-426">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-426">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-427">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-427">1.0</span></span>|
|[<span data-ttu-id="8e1c2-428">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-428">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-429">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-430">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-430">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-431">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-432">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-432">Example</span></span>

<span data-ttu-id="8e1c2-p119">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="8e1c2-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="8e1c2-436">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-436">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8e1c2-437">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-437">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-438">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-438">Type:</span></span>

*   [<span data-ttu-id="8e1c2-439">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8e1c2-439">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="8e1c2-440">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-440">Requirements</span></span>

|<span data-ttu-id="8e1c2-441">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-441">Requirement</span></span>|<span data-ttu-id="8e1c2-442">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-443">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-443">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-444">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-444">1.0</span></span>|
|[<span data-ttu-id="8e1c2-445">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-446">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-447">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-448">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-449">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-449">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="8e1c2-450">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-450">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="8e1c2-451">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-451">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e1c2-452">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-452">Read mode</span></span>

<span data-ttu-id="8e1c2-453">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-453">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e1c2-454">Mode composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-454">Compose mode</span></span>

<span data-ttu-id="8e1c2-455">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-455">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-456">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-456">Type:</span></span>

*   <span data-ttu-id="8e1c2-457">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-457">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-458">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-458">Requirements</span></span>

|<span data-ttu-id="8e1c2-459">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-459">Requirement</span></span>|<span data-ttu-id="8e1c2-460">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-461">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-461">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-462">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-462">1.0</span></span>|
|[<span data-ttu-id="8e1c2-463">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-464">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-465">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-466">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-466">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-467">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-467">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8e1c2-468">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-468">normalizedSubject :String</span></span>

<span data-ttu-id="8e1c2-p120">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8e1c2-p121">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-473">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-473">Type:</span></span>

*   <span data-ttu-id="8e1c2-474">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8e1c2-474">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-475">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-475">Requirements</span></span>

|<span data-ttu-id="8e1c2-476">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-476">Requirement</span></span>|<span data-ttu-id="8e1c2-477">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-477">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-478">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-478">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-479">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-479">1.0</span></span>|
|[<span data-ttu-id="8e1c2-480">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-480">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-481">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-481">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-482">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-482">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-483">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-483">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-484">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-484">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="8e1c2-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="8e1c2-486">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-486">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-487">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-487">Type:</span></span>

*   [<span data-ttu-id="8e1c2-488">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="8e1c2-488">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="8e1c2-489">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-489">Requirements</span></span>

|<span data-ttu-id="8e1c2-490">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-490">Requirement</span></span>|<span data-ttu-id="8e1c2-491">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-492">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-492">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-493">1.3</span><span class="sxs-lookup"><span data-stu-id="8e1c2-493">1.3</span></span>|
|[<span data-ttu-id="8e1c2-494">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-495">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-496">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-497">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-497">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8e1c2-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8e1c2-499">Fournit l’accès à un événement les participants facultatifs.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-499">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8e1c2-500">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-500">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e1c2-501">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-501">Read mode</span></span>

<span data-ttu-id="8e1c2-502">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-502">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e1c2-503">Mode composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-503">Compose mode</span></span>

<span data-ttu-id="8e1c2-504">Le `optionalAttendees` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs à une réunion.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-504">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-505">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-505">Type:</span></span>

*   <span data-ttu-id="8e1c2-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-507">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-507">Requirements</span></span>

|<span data-ttu-id="8e1c2-508">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-508">Requirement</span></span>|<span data-ttu-id="8e1c2-509">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-510">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-510">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-511">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-511">1.0</span></span>|
|[<span data-ttu-id="8e1c2-512">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-512">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-513">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-514">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-514">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-515">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-515">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-516">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-516">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="8e1c2-517">organisateur :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[organisateur](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-517">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="8e1c2-518">Obtient l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-518">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e1c2-519">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-519">Read mode</span></span>

<span data-ttu-id="8e1c2-520">Le `organizer` propriété renvoie un objet [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-520">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e1c2-521">Mode composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-521">Compose mode</span></span>

<span data-ttu-id="8e1c2-522">Le `organizer` propriété retourne un objet [organisateur](/javascript/api/outlook/office.organizer) qui fournit une méthode pour obtenir la valeur de l’organisateur.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-522">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-523">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-523">Type:</span></span>

*   <span data-ttu-id="8e1c2-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [organisateur](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-525">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-525">Requirements</span></span>

|<span data-ttu-id="8e1c2-526">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-526">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="8e1c2-527">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-527">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-528">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-528">1.0</span></span>|<span data-ttu-id="8e1c2-529">1.7</span><span class="sxs-lookup"><span data-stu-id="8e1c2-529">1.7</span></span>|
|[<span data-ttu-id="8e1c2-530">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-530">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-531">ReadItem</span></span>|<span data-ttu-id="8e1c2-532">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-532">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e1c2-533">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-534">Read</span><span class="sxs-lookup"><span data-stu-id="8e1c2-534">Read</span></span>|<span data-ttu-id="8e1c2-535">Composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-536">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-536">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="8e1c2-537">périodicité (nullable) :[périodicité](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-537">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="8e1c2-538">Obtient ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-538">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="8e1c2-539">Obtient la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-539">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="8e1c2-540">Lecture et de composition modes pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-540">Read and compose modes for appointment items.</span></span> <span data-ttu-id="8e1c2-541">Mode de lecture pour les éléments de la demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-541">Read mode for meeting request items.</span></span>

<span data-ttu-id="8e1c2-542">Le `recurrence` propriété renvoie un objet de [périodicité](/javascript/api/outlook/office.recurrence) pour les demandes de réunions ou de rendez-vous périodiques si un élément est une série ou à une instance d’une série.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-542">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="8e1c2-543">`null`sont retournées pour les rendez-vous uniques et des demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-543">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="8e1c2-544">`undefined`sont retournées pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-544">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="8e1c2-545">Remarque : Les demandes de réunion ont une `itemClass` valeur IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-545">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="8e1c2-546">Remarque : Si l’objet de périodicité est `null`, cela indique que l’objet est un rendez-vous ou une demande de réunion d’un rendez-vous et ne fait pas partie d’une série.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-546">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-547">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-547">Type:</span></span>

* [<span data-ttu-id="8e1c2-548">Périodicité</span><span class="sxs-lookup"><span data-stu-id="8e1c2-548">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="8e1c2-549">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-549">Requirement</span></span>|<span data-ttu-id="8e1c2-550">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-551">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-551">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-552">1.7</span><span class="sxs-lookup"><span data-stu-id="8e1c2-552">1.7</span></span>|
|[<span data-ttu-id="8e1c2-553">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-554">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-555">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-556">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-556">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8e1c2-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8e1c2-558">Fournit l’accès à des participants obligatoires d’un événement.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-558">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8e1c2-559">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-559">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e1c2-560">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-560">Read mode</span></span>

<span data-ttu-id="8e1c2-561">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-561">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e1c2-562">Mode composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-562">Compose mode</span></span>

<span data-ttu-id="8e1c2-563">Le `requiredAttendees` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les participants obligatoires à une réunion.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-563">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-564">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-564">Type:</span></span>

*   <span data-ttu-id="8e1c2-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-566">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-566">Requirements</span></span>

|<span data-ttu-id="8e1c2-567">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-567">Requirement</span></span>|<span data-ttu-id="8e1c2-568">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-568">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-569">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-569">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-570">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-570">1.0</span></span>|
|[<span data-ttu-id="8e1c2-571">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-572">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-573">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-573">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-574">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-574">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-575">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-575">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="8e1c2-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="8e1c2-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8e1c2-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-581">Le `recipientType` propriété de la `EmailAddressDetails` objet dans le `sender` propriété est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-582">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-582">Type:</span></span>

*   [<span data-ttu-id="8e1c2-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8e1c2-583">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8e1c2-584">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-584">Requirements</span></span>

|<span data-ttu-id="8e1c2-585">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-585">Requirement</span></span>|<span data-ttu-id="8e1c2-586">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-587">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-587">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-588">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-588">1.0</span></span>|
|[<span data-ttu-id="8e1c2-589">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-590">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-591">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-592">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-593">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-593">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="8e1c2-594">seriesId (nullable) : chaîne</span><span class="sxs-lookup"><span data-stu-id="8e1c2-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="8e1c2-595">Obtient l’id de la série appartenant à une instance.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="8e1c2-596">Dans OWA et Outlook, le `seriesId` renvoie l’ID Exchange Web Services (EWS) de l’élément parent (série) auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="8e1c2-597">Toutefois, dans iOS et Android, le `seriesId` renvoie l’ID du reste de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-598">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8e1c2-599">Le `seriesId` propriété n’est pas identique à l’ID Outlook utilisées par l’API REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="8e1c2-600">Avant d’effectuer des appels d’API REST à l’aide de cette valeur, elle doit être convertie à l’aide de [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8e1c2-601">Pour plus d’informations, voir [utiliser les API REST d’Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="8e1c2-602">Le `seriesId` propriété renvoie `null` pour les éléments qui n’ont pas d’éléments parents comme unique rendez-vous, les éléments de série, ou les demandes de réunion et renvoie `undefined` pour tous les éléments qui ne sont pas les demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-603">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-603">Type:</span></span>

* <span data-ttu-id="8e1c2-604">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8e1c2-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-605">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-605">Requirements</span></span>

|<span data-ttu-id="8e1c2-606">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-606">Requirement</span></span>|<span data-ttu-id="8e1c2-607">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-608">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-608">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-609">1.7</span><span class="sxs-lookup"><span data-stu-id="8e1c2-609">1.7</span></span>|
|[<span data-ttu-id="8e1c2-610">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-611">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-612">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-613">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-613">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-614">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-614">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="8e1c2-615">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-615">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="8e1c2-616">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8e1c2-p130">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e1c2-619">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-619">Read mode</span></span>

<span data-ttu-id="8e1c2-620">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-620">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e1c2-621">Mode composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-621">Compose mode</span></span>

<span data-ttu-id="8e1c2-622">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8e1c2-623">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-623">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-624">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-624">Type:</span></span>

*   <span data-ttu-id="8e1c2-625">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-625">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-626">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-626">Requirements</span></span>

|<span data-ttu-id="8e1c2-627">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-627">Requirement</span></span>|<span data-ttu-id="8e1c2-628">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-629">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-629">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-630">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-630">1.0</span></span>|
|[<span data-ttu-id="8e1c2-631">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-632">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-633">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-634">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-634">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-635">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-635">Example</span></span>

<span data-ttu-id="8e1c2-636">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-636">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="8e1c2-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="8e1c2-638">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-638">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8e1c2-639">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-639">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e1c2-640">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-640">Read mode</span></span>

<span data-ttu-id="8e1c2-p131">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="8e1c2-643">Mode composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-643">Compose mode</span></span>

<span data-ttu-id="8e1c2-644">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8e1c2-645">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-645">Type:</span></span>

*   <span data-ttu-id="8e1c2-646">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-646">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-647">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-647">Requirements</span></span>

|<span data-ttu-id="8e1c2-648">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-648">Requirement</span></span>|<span data-ttu-id="8e1c2-649">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-650">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-651">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-651">1.0</span></span>|
|[<span data-ttu-id="8e1c2-652">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-653">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-654">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-655">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-655">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8e1c2-656">à : Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8e1c2-657">Permet d’accéder aux destinataires de la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8e1c2-658">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e1c2-659">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-659">Read mode</span></span>

<span data-ttu-id="8e1c2-p133">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e1c2-662">Mode composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-662">Compose mode</span></span>

<span data-ttu-id="8e1c2-663">Le `to` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8e1c2-664">Type :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-664">Type:</span></span>

*   <span data-ttu-id="8e1c2-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-666">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-666">Requirements</span></span>

|<span data-ttu-id="8e1c2-667">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-667">Requirement</span></span>|<span data-ttu-id="8e1c2-668">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-669">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-669">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-670">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-670">1.0</span></span>|
|[<span data-ttu-id="8e1c2-671">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-672">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-673">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-674">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-674">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-675">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-675">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="8e1c2-676">Méthodes</span><span class="sxs-lookup"><span data-stu-id="8e1c2-676">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8e1c2-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8e1c2-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8e1c2-678">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-678">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8e1c2-679">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-679">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8e1c2-680">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-680">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-681">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-681">Parameters:</span></span>
|<span data-ttu-id="8e1c2-682">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-682">Name</span></span>|<span data-ttu-id="8e1c2-683">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-683">Type</span></span>|<span data-ttu-id="8e1c2-684">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-684">Attributes</span></span>|<span data-ttu-id="8e1c2-685">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-685">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="8e1c2-686">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-686">String</span></span>||<span data-ttu-id="8e1c2-p134">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="8e1c2-689">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-689">String</span></span>||<span data-ttu-id="8e1c2-p135">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8e1c2-692">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-692">Object</span></span>|<span data-ttu-id="8e1c2-693">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-693">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-694">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-694">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e1c2-695">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-695">Object</span></span>|<span data-ttu-id="8e1c2-696">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-696">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-697">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-697">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="8e1c2-698">Boolean</span><span class="sxs-lookup"><span data-stu-id="8e1c2-698">Boolean</span></span>|<span data-ttu-id="8e1c2-699">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-699">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-700">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-700">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="8e1c2-701">fonction</span><span class="sxs-lookup"><span data-stu-id="8e1c2-701">function</span></span>|<span data-ttu-id="8e1c2-702">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-702">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-703">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-703">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8e1c2-704">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-704">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8e1c2-705">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-705">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8e1c2-706">Erreurs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-706">Errors</span></span>

|<span data-ttu-id="8e1c2-707">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-707">Error code</span></span>|<span data-ttu-id="8e1c2-708">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-708">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="8e1c2-709">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-709">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="8e1c2-710">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-710">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8e1c2-711">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-711">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-712">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-712">Requirements</span></span>

|<span data-ttu-id="8e1c2-713">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-713">Requirement</span></span>|<span data-ttu-id="8e1c2-714">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-715">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-715">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-716">1.1</span><span class="sxs-lookup"><span data-stu-id="8e1c2-716">1.1</span></span>|
|[<span data-ttu-id="8e1c2-717">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-717">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-718">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-718">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e1c2-719">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-719">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-720">Composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-720">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8e1c2-721">範例</span><span class="sxs-lookup"><span data-stu-id="8e1c2-721">Examples</span></span>

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

<span data-ttu-id="8e1c2-722">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-722">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="8e1c2-723">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [rappel])</span><span class="sxs-lookup"><span data-stu-id="8e1c2-723">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8e1c2-724">Ajoute un fichier à partir de l’en base64 codage à un message ou un rendez-vous en tant que pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-724">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8e1c2-725">Le `addFileAttachmentFromBase64Async` méthode télécharge le fichier à partir du codage en base64 et l’attache à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-725">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="8e1c2-726">Cette méthode renvoie l’identificateur de pièce jointe dans l’objet AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-726">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="8e1c2-727">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-727">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-728">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-728">Parameters:</span></span>
|<span data-ttu-id="8e1c2-729">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-729">Name</span></span>|<span data-ttu-id="8e1c2-730">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-730">Type</span></span>|<span data-ttu-id="8e1c2-731">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-731">Attributes</span></span>|<span data-ttu-id="8e1c2-732">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-732">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="8e1c2-733">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-733">String</span></span>||<span data-ttu-id="8e1c2-734">Le contenu d’une image ou un fichier à ajouter à un courrier électronique ou un événement codage Base64.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-734">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="8e1c2-735">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-735">String</span></span>||<span data-ttu-id="8e1c2-p137">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8e1c2-738">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-738">Object</span></span>|<span data-ttu-id="8e1c2-739">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-739">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-740">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e1c2-741">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-741">Object</span></span>|<span data-ttu-id="8e1c2-742">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-742">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-743">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="8e1c2-744">Boolean</span><span class="sxs-lookup"><span data-stu-id="8e1c2-744">Boolean</span></span>|<span data-ttu-id="8e1c2-745">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-745">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-746">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="8e1c2-747">fonction</span><span class="sxs-lookup"><span data-stu-id="8e1c2-747">function</span></span>|<span data-ttu-id="8e1c2-748">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-748">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-749">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8e1c2-750">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8e1c2-751">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8e1c2-752">Erreurs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-752">Errors</span></span>

|<span data-ttu-id="8e1c2-753">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-753">Error code</span></span>|<span data-ttu-id="8e1c2-754">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="8e1c2-755">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="8e1c2-756">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8e1c2-757">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-758">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-758">Requirements</span></span>

|<span data-ttu-id="8e1c2-759">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-759">Requirement</span></span>|<span data-ttu-id="8e1c2-760">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-761">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-761">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-762">Aperçu</span><span class="sxs-lookup"><span data-stu-id="8e1c2-762">Preview</span></span>|
|[<span data-ttu-id="8e1c2-763">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e1c2-765">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-766">Composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8e1c2-767">範例</span><span class="sxs-lookup"><span data-stu-id="8e1c2-767">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="8e1c2-768">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8e1c2-768">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="8e1c2-769">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-769">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="8e1c2-770">Actuellement les types d’événements pris en charge sont `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, et`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="8e1c2-770">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-771">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-771">Parameters:</span></span>

| <span data-ttu-id="8e1c2-772">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-772">Name</span></span> | <span data-ttu-id="8e1c2-773">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-773">Type</span></span> | <span data-ttu-id="8e1c2-774">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-774">Attributes</span></span> | <span data-ttu-id="8e1c2-775">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-775">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="8e1c2-776">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="8e1c2-776">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="8e1c2-777">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-777">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="8e1c2-778">Fonction</span><span class="sxs-lookup"><span data-stu-id="8e1c2-778">Function</span></span> || <span data-ttu-id="8e1c2-p138">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="8e1c2-782">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-782">Object</span></span> | <span data-ttu-id="8e1c2-783">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-783">&lt;optional&gt;</span></span> | <span data-ttu-id="8e1c2-784">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-784">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8e1c2-785">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-785">Object</span></span> | <span data-ttu-id="8e1c2-786">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-786">&lt;optional&gt;</span></span> | <span data-ttu-id="8e1c2-787">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-787">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="8e1c2-788">fonction</span><span class="sxs-lookup"><span data-stu-id="8e1c2-788">function</span></span>| <span data-ttu-id="8e1c2-789">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-789">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-790">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-791">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-791">Requirements</span></span>

|<span data-ttu-id="8e1c2-792">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-792">Requirement</span></span>| <span data-ttu-id="8e1c2-793">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-793">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-794">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-794">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e1c2-795">1.7</span><span class="sxs-lookup"><span data-stu-id="8e1c2-795">1.7</span></span> |
|[<span data-ttu-id="8e1c2-796">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-796">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e1c2-797">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-797">ReadItem</span></span> |
|[<span data-ttu-id="8e1c2-798">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-798">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e1c2-799">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-799">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8e1c2-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8e1c2-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8e1c2-801">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-801">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8e1c2-p139">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8e1c2-805">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-805">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8e1c2-806">Si votre complément Office est en cours d’exécution dans Outlook Web App, le `addItemAttachmentAsync` méthode permettre joindre des éléments à des éléments autres que l’élément à modifier ; Toutefois, cela n’est pas pris en charge et n’est pas recommandée.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-806">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-807">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-807">Parameters:</span></span>

|<span data-ttu-id="8e1c2-808">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-808">Name</span></span>|<span data-ttu-id="8e1c2-809">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-809">Type</span></span>|<span data-ttu-id="8e1c2-810">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-810">Attributes</span></span>|<span data-ttu-id="8e1c2-811">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-811">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="8e1c2-812">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-812">String</span></span>||<span data-ttu-id="8e1c2-p140">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="8e1c2-815">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-815">String</span></span>||<span data-ttu-id="8e1c2-p141">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8e1c2-818">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-818">Object</span></span>|<span data-ttu-id="8e1c2-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-819">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-820">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-820">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e1c2-821">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-821">Object</span></span>|<span data-ttu-id="8e1c2-822">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-822">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-823">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-823">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e1c2-824">fonction</span><span class="sxs-lookup"><span data-stu-id="8e1c2-824">function</span></span>|<span data-ttu-id="8e1c2-825">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-825">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-826">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-826">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8e1c2-827">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-827">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8e1c2-828">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-828">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8e1c2-829">Erreurs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-829">Errors</span></span>

|<span data-ttu-id="8e1c2-830">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-830">Error code</span></span>|<span data-ttu-id="8e1c2-831">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-831">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8e1c2-832">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-832">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-833">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-833">Requirements</span></span>

|<span data-ttu-id="8e1c2-834">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-834">Requirement</span></span>|<span data-ttu-id="8e1c2-835">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-836">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-837">1.1</span><span class="sxs-lookup"><span data-stu-id="8e1c2-837">1.1</span></span>|
|[<span data-ttu-id="8e1c2-838">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-839">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-839">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e1c2-840">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-841">Composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-841">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-842">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-842">Example</span></span>

<span data-ttu-id="8e1c2-843">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-843">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="8e1c2-844">close()</span><span class="sxs-lookup"><span data-stu-id="8e1c2-844">close()</span></span>

<span data-ttu-id="8e1c2-845">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-845">Closes the current item that is being composed.</span></span>

<span data-ttu-id="8e1c2-p142">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-848">Dans Outlook sur le web, si l’élément est un rendez-vous et il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, supprimer ou annuler même si aucune modification de l’élément depuis le dernier enregistrée.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-848">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="8e1c2-849">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-849">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-850">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-850">Requirements</span></span>

|<span data-ttu-id="8e1c2-851">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-851">Requirement</span></span>|<span data-ttu-id="8e1c2-852">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-853">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-853">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-854">1.3</span><span class="sxs-lookup"><span data-stu-id="8e1c2-854">1.3</span></span>|
|[<span data-ttu-id="8e1c2-855">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-855">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-856">Restricted</span><span class="sxs-lookup"><span data-stu-id="8e1c2-856">Restricted</span></span>|
|[<span data-ttu-id="8e1c2-857">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-857">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-858">Composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-858">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="8e1c2-859">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-859">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="8e1c2-860">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-860">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-861">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-861">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e1c2-862">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-862">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8e1c2-863">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-863">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="8e1c2-p143">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-867">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-867">Parameters:</span></span>

|<span data-ttu-id="8e1c2-868">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-868">Name</span></span>|<span data-ttu-id="8e1c2-869">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-869">Type</span></span>|<span data-ttu-id="8e1c2-870">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-870">Attributes</span></span>|<span data-ttu-id="8e1c2-871">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-871">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="8e1c2-872">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8e1c2-872">String &#124; Object</span></span>||<span data-ttu-id="8e1c2-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8e1c2-875">**OU**</span><span class="sxs-lookup"><span data-stu-id="8e1c2-875">**OR**</span></span><br/><span data-ttu-id="8e1c2-p145">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="8e1c2-878">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-878">String</span></span>|<span data-ttu-id="8e1c2-879">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-879">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="8e1c2-882">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-882">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="8e1c2-883">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-883">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-884">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-884">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="8e1c2-885">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-885">String</span></span>||<span data-ttu-id="8e1c2-p147">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="8e1c2-888">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-888">String</span></span>||<span data-ttu-id="8e1c2-889">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-889">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="8e1c2-890">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-890">String</span></span>||<span data-ttu-id="8e1c2-p148">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="8e1c2-893">Boolean</span><span class="sxs-lookup"><span data-stu-id="8e1c2-893">Boolean</span></span>||<span data-ttu-id="8e1c2-p149">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="8e1c2-896">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-896">String</span></span>||<span data-ttu-id="8e1c2-p150">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="8e1c2-900">function</span><span class="sxs-lookup"><span data-stu-id="8e1c2-900">function</span></span>|<span data-ttu-id="8e1c2-901">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-901">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-902">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-902">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-903">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-903">Requirements</span></span>

|<span data-ttu-id="8e1c2-904">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-904">Requirement</span></span>|<span data-ttu-id="8e1c2-905">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-906">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-906">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-907">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-907">1.0</span></span>|
|[<span data-ttu-id="8e1c2-908">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-909">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-910">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-911">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-911">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8e1c2-912">Exemples</span><span class="sxs-lookup"><span data-stu-id="8e1c2-912">Examples</span></span>

<span data-ttu-id="8e1c2-913">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-913">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8e1c2-914">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-914">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8e1c2-915">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-915">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8e1c2-916">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-916">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8e1c2-917">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-917">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8e1c2-918">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-918">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="8e1c2-919">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-919">displayReplyForm(formData)</span></span>

<span data-ttu-id="8e1c2-920">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-920">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-921">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e1c2-922">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-922">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8e1c2-923">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-923">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="8e1c2-p151">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-927">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-927">Parameters:</span></span>

|<span data-ttu-id="8e1c2-928">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-928">Name</span></span>|<span data-ttu-id="8e1c2-929">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-929">Type</span></span>|<span data-ttu-id="8e1c2-930">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-930">Attributes</span></span>|<span data-ttu-id="8e1c2-931">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-931">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="8e1c2-932">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8e1c2-932">String &#124; Object</span></span>||<span data-ttu-id="8e1c2-p152">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8e1c2-935">**OU**</span><span class="sxs-lookup"><span data-stu-id="8e1c2-935">**OR**</span></span><br/><span data-ttu-id="8e1c2-p153">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="8e1c2-938">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-938">String</span></span>|<span data-ttu-id="8e1c2-939">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-939">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="8e1c2-942">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-942">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="8e1c2-943">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-943">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-944">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-944">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="8e1c2-945">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-945">String</span></span>||<span data-ttu-id="8e1c2-p155">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="8e1c2-948">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-948">String</span></span>||<span data-ttu-id="8e1c2-949">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-949">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="8e1c2-950">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-950">String</span></span>||<span data-ttu-id="8e1c2-p156">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="8e1c2-953">Boolean</span><span class="sxs-lookup"><span data-stu-id="8e1c2-953">Boolean</span></span>||<span data-ttu-id="8e1c2-p157">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="8e1c2-956">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-956">String</span></span>||<span data-ttu-id="8e1c2-p158">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="8e1c2-960">function</span><span class="sxs-lookup"><span data-stu-id="8e1c2-960">function</span></span>|<span data-ttu-id="8e1c2-961">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-961">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-962">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-962">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-963">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-963">Requirements</span></span>

|<span data-ttu-id="8e1c2-964">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-964">Requirement</span></span>|<span data-ttu-id="8e1c2-965">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-966">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-966">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-967">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-967">1.0</span></span>|
|[<span data-ttu-id="8e1c2-968">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-968">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-969">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-970">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-970">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-971">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-971">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8e1c2-972">Exemples</span><span class="sxs-lookup"><span data-stu-id="8e1c2-972">Examples</span></span>

<span data-ttu-id="8e1c2-973">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-973">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8e1c2-974">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-974">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8e1c2-975">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-975">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8e1c2-976">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-976">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8e1c2-977">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-977">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8e1c2-978">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-978">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="8e1c2-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8e1c2-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="8e1c2-980">Obtient les entités trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-980">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-981">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-981">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-982">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-982">Requirements</span></span>

|<span data-ttu-id="8e1c2-983">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-983">Requirement</span></span>|<span data-ttu-id="8e1c2-984">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-985">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-985">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-986">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-986">1.0</span></span>|
|[<span data-ttu-id="8e1c2-987">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-988">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-989">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-990">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-990">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e1c2-991">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-991">Returns:</span></span>

<span data-ttu-id="8e1c2-992">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-992">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8e1c2-993">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-993">Example</span></span>

<span data-ttu-id="8e1c2-994">L’exemple suivant accède aux entités de contacts dans le corps de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-994">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="8e1c2-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8e1c2-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8e1c2-996">Obtient un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-996">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-997">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-997">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-998">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-998">Parameters:</span></span>

|<span data-ttu-id="8e1c2-999">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-999">Name</span></span>|<span data-ttu-id="8e1c2-1000">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1000">Type</span></span>|<span data-ttu-id="8e1c2-1001">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1001">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="8e1c2-1002">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1002">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="8e1c2-1003">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1003">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1004">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1004">Requirements</span></span>

|<span data-ttu-id="8e1c2-1005">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1005">Requirement</span></span>|<span data-ttu-id="8e1c2-1006">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1007">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1007">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1008">1.0</span></span>|
|[<span data-ttu-id="8e1c2-1009">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-1010">Restricted</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1010">Restricted</span></span>|
|[<span data-ttu-id="8e1c2-1011">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-1012">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1012">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e1c2-1013">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1013">Returns:</span></span>

<span data-ttu-id="8e1c2-1014">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1014">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8e1c2-1015">Si aucune entité du type spécifié n’est présentes dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1015">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="8e1c2-1016">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1016">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8e1c2-1017">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1017">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="8e1c2-1018">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1018">Value of `entityType`</span></span>|<span data-ttu-id="8e1c2-1019">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1019">Type of objects in returned array</span></span>|<span data-ttu-id="8e1c2-1020">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1020">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="8e1c2-1021">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1021">String</span></span>|<span data-ttu-id="8e1c2-1022">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1022">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="8e1c2-1023">Contact</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1023">Contact</span></span>|<span data-ttu-id="8e1c2-1024">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1024">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="8e1c2-1025">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1025">String</span></span>|<span data-ttu-id="8e1c2-1026">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1026">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="8e1c2-1027">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1027">MeetingSuggestion</span></span>|<span data-ttu-id="8e1c2-1028">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1028">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="8e1c2-1029">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1029">PhoneNumber</span></span>|<span data-ttu-id="8e1c2-1030">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1030">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="8e1c2-1031">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1031">TaskSuggestion</span></span>|<span data-ttu-id="8e1c2-1032">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1032">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="8e1c2-1033">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1033">String</span></span>|<span data-ttu-id="8e1c2-1034">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1034">**Restricted**</span></span>|

<span data-ttu-id="8e1c2-1035">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8e1c2-1035">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="8e1c2-1036">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1036">Example</span></span>

<span data-ttu-id="8e1c2-1037">L’exemple suivant montre comment accéder à un tableau de chaînes qui représentent les adresses postales dans le corps de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1037">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="8e1c2-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8e1c2-1039">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1039">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-1040">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1040">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e1c2-1041">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1041">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-1042">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1042">Parameters:</span></span>

|<span data-ttu-id="8e1c2-1043">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1043">Name</span></span>|<span data-ttu-id="8e1c2-1044">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1044">Type</span></span>|<span data-ttu-id="8e1c2-1045">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1045">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="8e1c2-1046">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1046">String</span></span>|<span data-ttu-id="8e1c2-1047">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1047">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1048">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1048">Requirements</span></span>

|<span data-ttu-id="8e1c2-1049">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1049">Requirement</span></span>|<span data-ttu-id="8e1c2-1050">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1051">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1051">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-1052">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1052">1.0</span></span>|
|[<span data-ttu-id="8e1c2-1053">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1054">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-1055">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-1056">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1056">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e1c2-1057">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1057">Returns:</span></span>

<span data-ttu-id="8e1c2-p160">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="8e1c2-1060">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8e1c2-1060">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="8e1c2-1061">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1061">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="8e1c2-1062">Récupère les données d’initialisation transmises quand le complément est [activé par un message actionnable](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1062">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-1063">Cette méthode est uniquement pris en charge par Outlook 2016 ou version ultérieure pour Windows (versions Click-to-Run 16.0.8413.1000 au plus tard) et Outlook sur le site web d’Office 365.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1063">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-1064">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1064">Parameters:</span></span>
|<span data-ttu-id="8e1c2-1065">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1065">Name</span></span>|<span data-ttu-id="8e1c2-1066">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1066">Type</span></span>|<span data-ttu-id="8e1c2-1067">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1067">Attributes</span></span>|<span data-ttu-id="8e1c2-1068">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1068">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8e1c2-1069">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1069">Object</span></span>|<span data-ttu-id="8e1c2-1070">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1070">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1071">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1071">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e1c2-1072">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1072">Object</span></span>|<span data-ttu-id="8e1c2-1073">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1073">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1074">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1074">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e1c2-1075">fonction</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1075">function</span></span>|<span data-ttu-id="8e1c2-1076">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1076">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1077">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1077">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8e1c2-1078">En cas de réussite, les données d’initialisation sont fournies dans le `asyncResult.value` propriété sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1078">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="8e1c2-1079">S’il n’existe aucun contexte d’initialisation, l’objet `asyncResult` contient un objet `Error` dont la propriété `code` est définie sur `9020` et la propriété `name` sur `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1079">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1080">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1080">Requirements</span></span>

|<span data-ttu-id="8e1c2-1081">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1081">Requirement</span></span>|<span data-ttu-id="8e1c2-1082">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1083">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1083">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-1084">Aperçu</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1084">Preview</span></span>|
|[<span data-ttu-id="8e1c2-1085">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1085">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1086">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-1087">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1087">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-1088">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1088">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-1089">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1089">Example</span></span>

```
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="8e1c2-1090">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1090">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8e1c2-1091">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1091">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-1092">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1092">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e1c2-p161">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8e1c2-1096">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1096">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8e1c2-1097">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1097">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8e1c2-p162">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1101">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1101">Requirements</span></span>

|<span data-ttu-id="8e1c2-1102">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1102">Requirement</span></span>|<span data-ttu-id="8e1c2-1103">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1103">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1104">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1104">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-1105">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1105">1.0</span></span>|
|[<span data-ttu-id="8e1c2-1106">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1106">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-1107">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1107">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-1108">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1108">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-1109">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1109">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e1c2-1110">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1110">Returns:</span></span>

<span data-ttu-id="8e1c2-p163">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="8e1c2-1113">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1113">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8e1c2-1114">Object</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1114">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8e1c2-1115">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1115">Example</span></span>

<span data-ttu-id="8e1c2-1116">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1116">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8e1c2-1117">getregexmatchesbyname (Name) → (nullable) {tableau. < chaîne >}</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1117">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8e1c2-1118">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1118">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-1119">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1119">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e1c2-1120">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1120">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8e1c2-p164">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-1123">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1123">Parameters:</span></span>

|<span data-ttu-id="8e1c2-1124">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1124">Name</span></span>|<span data-ttu-id="8e1c2-1125">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1125">Type</span></span>|<span data-ttu-id="8e1c2-1126">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1126">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="8e1c2-1127">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1127">String</span></span>|<span data-ttu-id="8e1c2-1128">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1128">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1129">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1129">Requirements</span></span>

|<span data-ttu-id="8e1c2-1130">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1130">Requirement</span></span>|<span data-ttu-id="8e1c2-1131">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1132">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-1133">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1133">1.0</span></span>|
|[<span data-ttu-id="8e1c2-1134">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-1135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1135">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-1136">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-1137">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1137">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e1c2-1138">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1138">Returns:</span></span>

<span data-ttu-id="8e1c2-1139">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1139">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="8e1c2-1140">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1140">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8e1c2-1141">< Chaîne > tableau.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1141">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8e1c2-1142">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1142">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="8e1c2-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="8e1c2-1144">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1144">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="8e1c2-p165">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-1147">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1147">Parameters:</span></span>

|<span data-ttu-id="8e1c2-1148">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1148">Name</span></span>|<span data-ttu-id="8e1c2-1149">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1149">Type</span></span>|<span data-ttu-id="8e1c2-1150">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1150">Attributes</span></span>|<span data-ttu-id="8e1c2-1151">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1151">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="8e1c2-1152">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1152">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="8e1c2-p166">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="8e1c2-1156">Object</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1156">Object</span></span>|<span data-ttu-id="8e1c2-1157">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1158">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1158">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e1c2-1159">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1159">Object</span></span>|<span data-ttu-id="8e1c2-1160">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1161">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1161">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e1c2-1162">fonction</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1162">function</span></span>||<span data-ttu-id="8e1c2-1163">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1163">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8e1c2-1164">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1164">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="8e1c2-1165">Pour accéder à la propriété source provenant de la sélection, appelez `asyncResult.value.sourceProperty`, qui sera soit `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1165">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1166">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1166">Requirements</span></span>

|<span data-ttu-id="8e1c2-1167">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1167">Requirement</span></span>|<span data-ttu-id="8e1c2-1168">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1168">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1169">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1169">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-1170">1.2</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1170">1.2</span></span>|
|[<span data-ttu-id="8e1c2-1171">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1171">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-1172">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1172">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e1c2-1173">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1173">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-1174">Composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1174">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e1c2-1175">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1175">Returns:</span></span>

<span data-ttu-id="8e1c2-1176">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1176">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="8e1c2-1177">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1177">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8e1c2-1178">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1178">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8e1c2-1179">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1179">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="8e1c2-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="8e1c2-p168">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-1183">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1183">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1184">Requirements</span></span>

|<span data-ttu-id="8e1c2-1185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1185">Requirement</span></span>|<span data-ttu-id="8e1c2-1186">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1186">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1187">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-1188">1.6</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1188">1.6</span></span>|
|[<span data-ttu-id="8e1c2-1189">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1189">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-1190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1190">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-1191">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-1192">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1192">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e1c2-1193">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1193">Returns:</span></span>

<span data-ttu-id="8e1c2-1194">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1194">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8e1c2-1195">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1195">Example</span></span>

<span data-ttu-id="8e1c2-1196">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1196">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="8e1c2-1197">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1197">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="8e1c2-p169">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-1200">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1200">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e1c2-p170">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8e1c2-1204">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1204">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8e1c2-1205">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1205">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8e1c2-p171">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1209">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1209">Requirements</span></span>

|<span data-ttu-id="8e1c2-1210">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1210">Requirement</span></span>|<span data-ttu-id="8e1c2-1211">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1212">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1212">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-1213">1.6</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1213">1.6</span></span>|
|[<span data-ttu-id="8e1c2-1214">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1214">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1215">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-1216">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-1217">Lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1217">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e1c2-1218">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1218">Returns:</span></span>

<span data-ttu-id="8e1c2-p172">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="8e1c2-1221">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1221">Example</span></span>

<span data-ttu-id="8e1c2-1222">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1222">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="8e1c2-1223">getSharedPropertiesAsync ([options], rappel)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1223">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="8e1c2-1224">Obtient les propriétés du rendez-vous sélectionné ou du message dans un dossier partagé, le calendrier ou la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1224">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-1225">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1225">Parameters:</span></span>

|<span data-ttu-id="8e1c2-1226">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1226">Name</span></span>|<span data-ttu-id="8e1c2-1227">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1227">Type</span></span>|<span data-ttu-id="8e1c2-1228">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1228">Attributes</span></span>|<span data-ttu-id="8e1c2-1229">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1229">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8e1c2-1230">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1230">Object</span></span>|<span data-ttu-id="8e1c2-1231">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1232">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1232">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e1c2-1233">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1233">Object</span></span>|<span data-ttu-id="8e1c2-1234">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1234">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1235">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1235">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e1c2-1236">fonction</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1236">function</span></span>||<span data-ttu-id="8e1c2-1237">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1237">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8e1c2-1238">Les propriétés partagées sont fournies en tant qu’un [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) objet dans le `asyncResult.value` propriété.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1238">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8e1c2-1239">Cet objet peut être utilisé pour obtenir les propriétés de l’élément partagé.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1239">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1240">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1240">Requirements</span></span>

|<span data-ttu-id="8e1c2-1241">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1241">Requirement</span></span>|<span data-ttu-id="8e1c2-1242">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1243">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1243">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-1244">Aperçu</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1244">Preview</span></span>|
|[<span data-ttu-id="8e1c2-1245">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1246">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-1247">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-1248">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-1249">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1249">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8e1c2-1250">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1250">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8e1c2-1251">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1251">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8e1c2-p174">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p174">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-1255">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1255">Parameters:</span></span>

|<span data-ttu-id="8e1c2-1256">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1256">Name</span></span>|<span data-ttu-id="8e1c2-1257">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1257">Type</span></span>|<span data-ttu-id="8e1c2-1258">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1258">Attributes</span></span>|<span data-ttu-id="8e1c2-1259">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1259">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="8e1c2-1260">function</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1260">function</span></span>||<span data-ttu-id="8e1c2-1261">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8e1c2-1262">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1262">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8e1c2-1263">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées de l’élément et enregistrer les modifications à la propriété personnalisée est redéfinie sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1263">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="8e1c2-1264">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1264">Object</span></span>|<span data-ttu-id="8e1c2-1265">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1266">Les développeurs peuvent fournir n’importe quel objet qu’ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1266">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="8e1c2-1267">Cet objet est accessible par le `asyncResult.asyncContext` propriété dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1267">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1268">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1268">Requirements</span></span>

|<span data-ttu-id="8e1c2-1269">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1269">Requirement</span></span>|<span data-ttu-id="8e1c2-1270">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1271">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1271">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1272">1.0</span></span>|
|[<span data-ttu-id="8e1c2-1273">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1274">ReadItem</span></span>|
|[<span data-ttu-id="8e1c2-1275">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-1276">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-1277">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1277">Example</span></span>

<span data-ttu-id="8e1c2-p177">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p177">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8e1c2-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8e1c2-1282">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1282">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8e1c2-p178">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p178">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-1287">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1287">Parameters:</span></span>

|<span data-ttu-id="8e1c2-1288">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1288">Name</span></span>|<span data-ttu-id="8e1c2-1289">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1289">Type</span></span>|<span data-ttu-id="8e1c2-1290">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1290">Attributes</span></span>|<span data-ttu-id="8e1c2-1291">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1291">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="8e1c2-1292">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1292">String</span></span>||<span data-ttu-id="8e1c2-p179">Identificateur de la pièce jointe à supprimer. La longueur maximale de la chaîne est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p179">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="8e1c2-1295">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1295">Object</span></span>|<span data-ttu-id="8e1c2-1296">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1297">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1297">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e1c2-1298">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1298">Object</span></span>|<span data-ttu-id="8e1c2-1299">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1299">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1300">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1300">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e1c2-1301">fonction</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1301">function</span></span>|<span data-ttu-id="8e1c2-1302">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1303">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1303">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8e1c2-1304">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1304">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8e1c2-1305">Erreurs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1305">Errors</span></span>

|<span data-ttu-id="8e1c2-1306">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1306">Error code</span></span>|<span data-ttu-id="8e1c2-1307">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1307">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="8e1c2-1308">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1308">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1309">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1309">Requirements</span></span>

|<span data-ttu-id="8e1c2-1310">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1310">Requirement</span></span>|<span data-ttu-id="8e1c2-1311">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1311">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1312">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1312">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-1313">1.1</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1313">1.1</span></span>|
|[<span data-ttu-id="8e1c2-1314">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-1315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1315">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e1c2-1316">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-1317">Composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1317">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-1318">Exemple</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1318">Example</span></span>

<span data-ttu-id="8e1c2-1319">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1319">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="8e1c2-1320">removeHandlerAsync (eventType, gestionnaire, [options], [rappel])</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1320">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="8e1c2-1321">Supprime un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1321">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="8e1c2-1322">Actuellement les types d’événements pris en charge sont `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, et`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1322">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-1323">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1323">Parameters:</span></span>

| <span data-ttu-id="8e1c2-1324">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1324">Name</span></span> | <span data-ttu-id="8e1c2-1325">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1325">Type</span></span> | <span data-ttu-id="8e1c2-1326">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1326">Attributes</span></span> | <span data-ttu-id="8e1c2-1327">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1327">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="8e1c2-1328">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1328">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="8e1c2-1329">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1329">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="8e1c2-1330">Fonction</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1330">Function</span></span> || <span data-ttu-id="8e1c2-p180">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p180">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="8e1c2-1334">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1334">Object</span></span> | <span data-ttu-id="8e1c2-1335">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1335">&lt;optional&gt;</span></span> | <span data-ttu-id="8e1c2-1336">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1336">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8e1c2-1337">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1337">Object</span></span> | <span data-ttu-id="8e1c2-1338">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1338">&lt;optional&gt;</span></span> | <span data-ttu-id="8e1c2-1339">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1339">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="8e1c2-1340">fonction</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1340">function</span></span>| <span data-ttu-id="8e1c2-1341">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1342">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1343">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1343">Requirements</span></span>

|<span data-ttu-id="8e1c2-1344">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1344">Requirement</span></span>| <span data-ttu-id="8e1c2-1345">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1346">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1346">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e1c2-1347">1.7</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1347">1.7</span></span> |
|[<span data-ttu-id="8e1c2-1348">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1348">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e1c2-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1349">ReadItem</span></span> |
|[<span data-ttu-id="8e1c2-1350">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1350">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e1c2-1351">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1351">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="8e1c2-1352">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1352">saveAsync([options], callback)</span></span>

<span data-ttu-id="8e1c2-1353">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1353">Asynchronously saves an item.</span></span>

<span data-ttu-id="8e1c2-p181">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-1357">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` pour utiliser avec EWS ou l’API REST, gardez à l’esprit que quand Outlook est en mode mis en cache, il peut prendre un certain temps avant que l’élément est réellement synchronisé avec le serveur.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1357">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="8e1c2-1358">Jusqu'à ce que l’élément est synchronisé, à l’aide de la `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1358">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="8e1c2-p183">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="8e1c2-1362">Les clients suivants ont un comportement différent pour `saveAsync` de rendez-vous dans le mode de composition :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1362">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="8e1c2-1363">Mac Outlook ne gère pas `saveAsync` d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1363">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="8e1c2-1364">Appel de `saveAsync` d’une réunion dans Outlook Mac renverra une erreur.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1364">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="8e1c2-1365">Outlook sur le web toujours envoie une invitation ou mettre à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1365">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-1366">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1366">Parameters:</span></span>

|<span data-ttu-id="8e1c2-1367">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1367">Name</span></span>|<span data-ttu-id="8e1c2-1368">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1368">Type</span></span>|<span data-ttu-id="8e1c2-1369">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1369">Attributes</span></span>|<span data-ttu-id="8e1c2-1370">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1370">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8e1c2-1371">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1371">Object</span></span>|<span data-ttu-id="8e1c2-1372">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1372">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1373">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1373">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e1c2-1374">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1374">Object</span></span>|<span data-ttu-id="8e1c2-1375">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1375">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1376">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1376">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e1c2-1377">fonction</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1377">function</span></span>||<span data-ttu-id="8e1c2-1378">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1378">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8e1c2-1379">En cas de réussite, l’identificateur d’élément est fournie dans le `asyncResult.value` propriété.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1379">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1380">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1380">Requirements</span></span>

|<span data-ttu-id="8e1c2-1381">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1381">Requirement</span></span>|<span data-ttu-id="8e1c2-1382">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1382">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1383">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1383">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-1384">1.3</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1384">1.3</span></span>|
|[<span data-ttu-id="8e1c2-1385">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-1386">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1386">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e1c2-1387">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-1388">Composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1388">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8e1c2-1389">範例</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1389">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="8e1c2-p185">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="8e1c2-1392">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1392">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="8e1c2-1393">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1393">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="8e1c2-p186">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e1c2-1397">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1397">Parameters:</span></span>

|<span data-ttu-id="8e1c2-1398">Nom</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1398">Name</span></span>|<span data-ttu-id="8e1c2-1399">Type</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1399">Type</span></span>|<span data-ttu-id="8e1c2-1400">Attributs</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1400">Attributes</span></span>|<span data-ttu-id="8e1c2-1401">Description</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1401">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="8e1c2-1402">String</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1402">String</span></span>||<span data-ttu-id="8e1c2-p187">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="8e1c2-1406">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1406">Object</span></span>|<span data-ttu-id="8e1c2-1407">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1407">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1408">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1408">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e1c2-1409">Objet</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1409">Object</span></span>|<span data-ttu-id="8e1c2-1410">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1410">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-1411">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1411">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="8e1c2-1412">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1412">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="8e1c2-1413">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1413">&lt;optional&gt;</span></span>|<span data-ttu-id="8e1c2-p188">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="8e1c2-p189">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="8e1c2-1418">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1418">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="8e1c2-1419">fonction</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1419">function</span></span>||<span data-ttu-id="8e1c2-1420">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1420">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e1c2-1421">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1421">Requirements</span></span>

|<span data-ttu-id="8e1c2-1422">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1422">Requirement</span></span>|<span data-ttu-id="8e1c2-1423">Valeur</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1423">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e1c2-1424">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e1c2-1425">1.2</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1425">1.2</span></span>|
|[<span data-ttu-id="8e1c2-1426">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e1c2-1427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1427">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e1c2-1428">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e1c2-1429">Composition</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1429">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8e1c2-1430">Exemples</span><span class="sxs-lookup"><span data-stu-id="8e1c2-1430">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```