
# <a name="item"></a><span data-ttu-id="a705f-101">élément</span><span class="sxs-lookup"><span data-stu-id="a705f-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="a705f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="a705f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="a705f-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="a705f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-105">Requirements</span></span>

|<span data-ttu-id="a705f-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-106">Requirement</span></span>|<span data-ttu-id="a705f-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-109">1.0</span></span>|
|[<span data-ttu-id="a705f-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-111">Restricted</span><span class="sxs-lookup"><span data-stu-id="a705f-111">Restricted</span></span>|
|[<span data-ttu-id="a705f-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a705f-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="a705f-114">Members and methods</span></span>

| <span data-ttu-id="a705f-115">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-115">Member</span></span> | <span data-ttu-id="a705f-116">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a705f-117">attachments</span><span class="sxs-lookup"><span data-stu-id="a705f-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="a705f-118">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-118">Member</span></span> |
| [<span data-ttu-id="a705f-119">bcc</span><span class="sxs-lookup"><span data-stu-id="a705f-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a705f-120">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-120">Member</span></span> |
| [<span data-ttu-id="a705f-121">body</span><span class="sxs-lookup"><span data-stu-id="a705f-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="a705f-122">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-122">Member</span></span> |
| [<span data-ttu-id="a705f-123">cc</span><span class="sxs-lookup"><span data-stu-id="a705f-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a705f-124">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-124">Member</span></span> |
| [<span data-ttu-id="a705f-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="a705f-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="a705f-126">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-126">Member</span></span> |
| [<span data-ttu-id="a705f-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="a705f-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="a705f-128">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-128">Member</span></span> |
| [<span data-ttu-id="a705f-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="a705f-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="a705f-130">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-130">Member</span></span> |
| [<span data-ttu-id="a705f-131">end</span><span class="sxs-lookup"><span data-stu-id="a705f-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="a705f-132">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-132">Member</span></span> |
| [<span data-ttu-id="a705f-133">from</span><span class="sxs-lookup"><span data-stu-id="a705f-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="a705f-134">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-134">Member</span></span> |
| [<span data-ttu-id="a705f-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="a705f-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="a705f-136">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-136">Member</span></span> |
| [<span data-ttu-id="a705f-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="a705f-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="a705f-138">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-138">Member</span></span> |
| [<span data-ttu-id="a705f-139">itemId</span><span class="sxs-lookup"><span data-stu-id="a705f-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="a705f-140">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-140">Member</span></span> |
| [<span data-ttu-id="a705f-141">itemType</span><span class="sxs-lookup"><span data-stu-id="a705f-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="a705f-142">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-142">Member</span></span> |
| [<span data-ttu-id="a705f-143">location</span><span class="sxs-lookup"><span data-stu-id="a705f-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="a705f-144">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-144">Member</span></span> |
| [<span data-ttu-id="a705f-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="a705f-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="a705f-146">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-146">Member</span></span> |
| [<span data-ttu-id="a705f-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="a705f-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="a705f-148">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-148">Member</span></span> |
| [<span data-ttu-id="a705f-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="a705f-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a705f-150">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-150">Member</span></span> |
| [<span data-ttu-id="a705f-151">organizer</span><span class="sxs-lookup"><span data-stu-id="a705f-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="a705f-152">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-152">Member</span></span> |
| [<span data-ttu-id="a705f-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="a705f-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="a705f-154">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-154">Member</span></span> |
| [<span data-ttu-id="a705f-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="a705f-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a705f-156">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-156">Member</span></span> |
| [<span data-ttu-id="a705f-157">sender</span><span class="sxs-lookup"><span data-stu-id="a705f-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="a705f-158">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-158">Member</span></span> |
| [<span data-ttu-id="a705f-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="a705f-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="a705f-160">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-160">Member</span></span> |
| [<span data-ttu-id="a705f-161">start</span><span class="sxs-lookup"><span data-stu-id="a705f-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="a705f-162">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-162">Member</span></span> |
| [<span data-ttu-id="a705f-163">subject</span><span class="sxs-lookup"><span data-stu-id="a705f-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="a705f-164">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-164">Member</span></span> |
| [<span data-ttu-id="a705f-165">to</span><span class="sxs-lookup"><span data-stu-id="a705f-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="a705f-166">Membre</span><span class="sxs-lookup"><span data-stu-id="a705f-166">Member</span></span> |
| [<span data-ttu-id="a705f-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a705f-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="a705f-168">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-168">Method</span></span> |
| [<span data-ttu-id="a705f-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="a705f-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="a705f-170">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-170">Method</span></span> |
| [<span data-ttu-id="a705f-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a705f-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a705f-172">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-172">Method</span></span> |
| [<span data-ttu-id="a705f-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a705f-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="a705f-174">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-174">Method</span></span> |
| [<span data-ttu-id="a705f-175">close</span><span class="sxs-lookup"><span data-stu-id="a705f-175">close</span></span>](#close) | <span data-ttu-id="a705f-176">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-176">Method</span></span> |
| [<span data-ttu-id="a705f-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="a705f-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="a705f-178">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-178">Method</span></span> |
| [<span data-ttu-id="a705f-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="a705f-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="a705f-180">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-180">Method</span></span> |
| [<span data-ttu-id="a705f-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="a705f-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="a705f-182">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-182">Method</span></span> |
| [<span data-ttu-id="a705f-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="a705f-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="a705f-184">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-184">Method</span></span> |
| [<span data-ttu-id="a705f-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="a705f-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="a705f-186">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-186">Method</span></span> |
| [<span data-ttu-id="a705f-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="a705f-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="a705f-188">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-188">Method</span></span> |
| [<span data-ttu-id="a705f-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a705f-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="a705f-190">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-190">Method</span></span> |
| [<span data-ttu-id="a705f-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="a705f-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="a705f-192">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-192">Method</span></span> |
| [<span data-ttu-id="a705f-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a705f-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="a705f-194">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-194">Method</span></span> |
| [<span data-ttu-id="a705f-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="a705f-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="a705f-196">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-196">Method</span></span> |
| [<span data-ttu-id="a705f-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a705f-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="a705f-198">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-198">Method</span></span> |
| [<span data-ttu-id="a705f-199">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a705f-199">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="a705f-200">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-200">Method</span></span> |
| [<span data-ttu-id="a705f-201">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a705f-201">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="a705f-202">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-202">Method</span></span> |
| [<span data-ttu-id="a705f-203">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a705f-203">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a705f-204">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-204">Method</span></span> |
| [<span data-ttu-id="a705f-205">saveAsync</span><span class="sxs-lookup"><span data-stu-id="a705f-205">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="a705f-206">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-206">Method</span></span> |
| [<span data-ttu-id="a705f-207">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a705f-207">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="a705f-208">Méthode</span><span class="sxs-lookup"><span data-stu-id="a705f-208">Method</span></span> |

### <a name="example"></a><span data-ttu-id="a705f-209">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-209">Example</span></span>

<span data-ttu-id="a705f-210">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="a705f-210">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="a705f-211">Membres</span><span class="sxs-lookup"><span data-stu-id="a705f-211">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="a705f-212">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a705f-212">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="a705f-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a705f-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-215">Certains types de fichiers bloqués par Outlook en raison de problèmes de sécurité potentiels et ne sont donc pas retournés.</span><span class="sxs-lookup"><span data-stu-id="a705f-215">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="a705f-216">Pour plus d’informations, voir les [pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="a705f-216">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-217">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-217">Type:</span></span>

*   <span data-ttu-id="a705f-218">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a705f-218">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-219">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-219">Requirements</span></span>

|<span data-ttu-id="a705f-220">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-220">Requirement</span></span>|<span data-ttu-id="a705f-221">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-222">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-222">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-223">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-223">1.0</span></span>|
|[<span data-ttu-id="a705f-224">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-225">ReadItem</span></span>|
|[<span data-ttu-id="a705f-226">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-227">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-227">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-228">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-228">Example</span></span>

<span data-ttu-id="a705f-229">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="a705f-229">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a705f-230">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a705f-230">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a705f-231">Obtient un objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires des Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="a705f-231">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="a705f-232">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="a705f-232">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-233">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-233">Type:</span></span>

*   [<span data-ttu-id="a705f-234">Destinataires</span><span class="sxs-lookup"><span data-stu-id="a705f-234">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="a705f-235">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-235">Requirements</span></span>

|<span data-ttu-id="a705f-236">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-236">Requirement</span></span>|<span data-ttu-id="a705f-237">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-238">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-238">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-239">1.1</span><span class="sxs-lookup"><span data-stu-id="a705f-239">1.1</span></span>|
|[<span data-ttu-id="a705f-240">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-240">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-241">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-241">ReadItem</span></span>|
|[<span data-ttu-id="a705f-242">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-242">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-243">Composition</span><span class="sxs-lookup"><span data-stu-id="a705f-243">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-244">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-244">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="a705f-245">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="a705f-245">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="a705f-246">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="a705f-246">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-247">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-247">Type:</span></span>

*   [<span data-ttu-id="a705f-248">Corps</span><span class="sxs-lookup"><span data-stu-id="a705f-248">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="a705f-249">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-249">Requirements</span></span>

|<span data-ttu-id="a705f-250">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-250">Requirement</span></span>|<span data-ttu-id="a705f-251">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-251">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-252">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-252">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-253">1.1</span><span class="sxs-lookup"><span data-stu-id="a705f-253">1.1</span></span>|
|[<span data-ttu-id="a705f-254">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-254">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-255">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-255">ReadItem</span></span>|
|[<span data-ttu-id="a705f-256">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-256">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-257">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-257">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a705f-258">cc : tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a705f-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a705f-259">Permet d’accéder aux destinataires Cc (copie carbone) d’un message.</span><span class="sxs-lookup"><span data-stu-id="a705f-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="a705f-260">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="a705f-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a705f-261">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-261">Read mode</span></span>

<span data-ttu-id="a705f-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="a705f-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a705f-264">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a705f-264">Compose mode</span></span>

<span data-ttu-id="a705f-265">Le `cc` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="a705f-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-266">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-266">Type:</span></span>

*   <span data-ttu-id="a705f-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a705f-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-268">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-268">Requirements</span></span>

|<span data-ttu-id="a705f-269">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-269">Requirement</span></span>|<span data-ttu-id="a705f-270">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-271">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-271">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-272">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-272">1.0</span></span>|
|[<span data-ttu-id="a705f-273">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-274">ReadItem</span></span>|
|[<span data-ttu-id="a705f-275">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-276">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-277">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-277">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="a705f-278">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="a705f-278">(nullable) conversationId :String</span></span>

<span data-ttu-id="a705f-279">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="a705f-279">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="a705f-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="a705f-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="a705f-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="a705f-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-284">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-284">Type:</span></span>

*   <span data-ttu-id="a705f-285">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a705f-285">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-286">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-286">Requirements</span></span>

|<span data-ttu-id="a705f-287">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-287">Requirement</span></span>|<span data-ttu-id="a705f-288">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-289">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-289">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-290">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-290">1.0</span></span>|
|[<span data-ttu-id="a705f-291">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-292">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-292">ReadItem</span></span>|
|[<span data-ttu-id="a705f-293">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-294">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-294">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="a705f-295">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="a705f-295">dateTimeCreated :Date</span></span>

<span data-ttu-id="a705f-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a705f-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-298">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-298">Type:</span></span>

*   <span data-ttu-id="a705f-299">Date</span><span class="sxs-lookup"><span data-stu-id="a705f-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-300">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-300">Requirements</span></span>

|<span data-ttu-id="a705f-301">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-301">Requirement</span></span>|<span data-ttu-id="a705f-302">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-303">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-304">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-304">1.0</span></span>|
|[<span data-ttu-id="a705f-305">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-306">ReadItem</span></span>|
|[<span data-ttu-id="a705f-307">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-308">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-309">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-309">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="a705f-310">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="a705f-310">dateTimeModified :Date</span></span>

<span data-ttu-id="a705f-p110">Obtient la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a705f-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-313">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a705f-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-314">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-314">Type:</span></span>

*   <span data-ttu-id="a705f-315">Date</span><span class="sxs-lookup"><span data-stu-id="a705f-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-316">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-316">Requirements</span></span>

|<span data-ttu-id="a705f-317">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-317">Requirement</span></span>|<span data-ttu-id="a705f-318">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-319">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-319">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-320">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-320">1.0</span></span>|
|[<span data-ttu-id="a705f-321">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-322">ReadItem</span></span>|
|[<span data-ttu-id="a705f-323">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-324">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-325">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-325">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="a705f-326">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a705f-326">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="a705f-327">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a705f-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="a705f-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="a705f-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a705f-330">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-330">Read mode</span></span>

<span data-ttu-id="a705f-331">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="a705f-331">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a705f-332">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a705f-332">Compose mode</span></span>

<span data-ttu-id="a705f-333">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="a705f-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="a705f-334">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="a705f-334">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-335">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-335">Type:</span></span>

*   <span data-ttu-id="a705f-336">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a705f-336">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-337">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-337">Requirements</span></span>

|<span data-ttu-id="a705f-338">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-338">Requirement</span></span>|<span data-ttu-id="a705f-339">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-340">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-341">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-341">1.0</span></span>|
|[<span data-ttu-id="a705f-342">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-343">ReadItem</span></span>|
|[<span data-ttu-id="a705f-344">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-345">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-346">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-346">Example</span></span>

<span data-ttu-id="a705f-347">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="a705f-347">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="a705f-348">à partir de :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[à partir de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="a705f-348">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="a705f-349">Obtient l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="a705f-349">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="a705f-p112">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="a705f-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-352">Le `recipientType` propriété de la `EmailAddressDetails` objet dans le `from` propriété est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a705f-352">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a705f-353">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-353">Read mode</span></span>

<span data-ttu-id="a705f-354">Le `from` propriété retourne un `EmailAddressDetails` objet.</span><span class="sxs-lookup"><span data-stu-id="a705f-354">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="a705f-355">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a705f-355">Compose mode</span></span>

<span data-ttu-id="a705f-356">Le `from` propriété renvoie une `From` objet qui fournit une méthode pour obtenir la valeur affectée à.</span><span class="sxs-lookup"><span data-stu-id="a705f-356">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a705f-357">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-357">Type:</span></span>

*   <span data-ttu-id="a705f-358">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [à partir de](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="a705f-358">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-359">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-359">Requirements</span></span>

|<span data-ttu-id="a705f-360">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-360">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a705f-361">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-362">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-362">1.0</span></span>|<span data-ttu-id="a705f-363">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a705f-363">Preview</span></span>|
|[<span data-ttu-id="a705f-364">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-365">ReadItem</span></span>|<span data-ttu-id="a705f-366">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a705f-366">ReadWriteItem</span></span>|
|[<span data-ttu-id="a705f-367">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-367">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-368">Read</span><span class="sxs-lookup"><span data-stu-id="a705f-368">Read</span></span>|<span data-ttu-id="a705f-369">Composition</span><span class="sxs-lookup"><span data-stu-id="a705f-369">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="a705f-370">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="a705f-370">internetMessageId :String</span></span>

<span data-ttu-id="a705f-p113">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a705f-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-373">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-373">Type:</span></span>

*   <span data-ttu-id="a705f-374">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a705f-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-375">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-375">Requirements</span></span>

|<span data-ttu-id="a705f-376">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-376">Requirement</span></span>|<span data-ttu-id="a705f-377">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-378">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-378">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-379">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-379">1.0</span></span>|
|[<span data-ttu-id="a705f-380">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-380">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-381">ReadItem</span></span>|
|[<span data-ttu-id="a705f-382">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-382">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-383">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-384">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-384">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="a705f-385">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="a705f-385">itemClass :String</span></span>

<span data-ttu-id="a705f-p114">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a705f-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="a705f-p115">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a705f-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="a705f-390">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-390">Type</span></span>|<span data-ttu-id="a705f-391">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-391">Description</span></span>|<span data-ttu-id="a705f-392">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="a705f-392">item class</span></span>|
|---|---|---|
|<span data-ttu-id="a705f-393">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="a705f-393">Appointment items</span></span>|<span data-ttu-id="a705f-394">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="a705f-394">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="a705f-395">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="a705f-395">Message items</span></span>|<span data-ttu-id="a705f-396">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="a705f-396">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="a705f-397">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="a705f-397">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-398">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-398">Type:</span></span>

*   <span data-ttu-id="a705f-399">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a705f-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-400">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-400">Requirements</span></span>

|<span data-ttu-id="a705f-401">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-401">Requirement</span></span>|<span data-ttu-id="a705f-402">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-403">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-403">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-404">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-404">1.0</span></span>|
|[<span data-ttu-id="a705f-405">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-406">ReadItem</span></span>|
|[<span data-ttu-id="a705f-407">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-408">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-409">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-409">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="a705f-410">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="a705f-410">(nullable) itemId :String</span></span>

<span data-ttu-id="a705f-p116">Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a705f-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-413">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="a705f-413">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a705f-414">Le `itemId` propriété n’est pas identique à l’identificateur d’entrée Outlook ou l’ID utilisé par l’API REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="a705f-414">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="a705f-415">Avant d’effectuer des appels d’API REST à l’aide de cette valeur, elle doit être convertie à l’aide de [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a705f-415">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a705f-416">Pour plus d’informations, voir [utiliser les API REST d’Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="a705f-416">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="a705f-p118">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-419">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-419">Type:</span></span>

*   <span data-ttu-id="a705f-420">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a705f-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-421">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-421">Requirements</span></span>

|<span data-ttu-id="a705f-422">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-422">Requirement</span></span>|<span data-ttu-id="a705f-423">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-424">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-425">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-425">1.0</span></span>|
|[<span data-ttu-id="a705f-426">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-427">ReadItem</span></span>|
|[<span data-ttu-id="a705f-428">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-429">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-430">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-430">Example</span></span>

<span data-ttu-id="a705f-p119">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a705f-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="a705f-433">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="a705f-433">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="a705f-434">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="a705f-434">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="a705f-435">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a705f-435">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-436">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-436">Type:</span></span>

*   [<span data-ttu-id="a705f-437">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="a705f-437">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="a705f-438">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-438">Requirements</span></span>

|<span data-ttu-id="a705f-439">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-439">Requirement</span></span>|<span data-ttu-id="a705f-440">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-440">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-441">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-441">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-442">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-442">1.0</span></span>|
|[<span data-ttu-id="a705f-443">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-443">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-444">ReadItem</span></span>|
|[<span data-ttu-id="a705f-445">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-445">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-446">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-446">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-447">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-447">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="a705f-448">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="a705f-448">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="a705f-449">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a705f-449">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a705f-450">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-450">Read mode</span></span>

<span data-ttu-id="a705f-451">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a705f-451">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a705f-452">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a705f-452">Compose mode</span></span>

<span data-ttu-id="a705f-453">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a705f-453">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-454">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-454">Type:</span></span>

*   <span data-ttu-id="a705f-455">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="a705f-455">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-456">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-456">Requirements</span></span>

|<span data-ttu-id="a705f-457">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-457">Requirement</span></span>|<span data-ttu-id="a705f-458">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-459">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-459">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-460">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-460">1.0</span></span>|
|[<span data-ttu-id="a705f-461">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-462">ReadItem</span></span>|
|[<span data-ttu-id="a705f-463">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-464">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-464">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-465">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-465">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="a705f-466">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="a705f-466">normalizedSubject :String</span></span>

<span data-ttu-id="a705f-p120">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a705f-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="a705f-p121">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="a705f-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-471">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-471">Type:</span></span>

*   <span data-ttu-id="a705f-472">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a705f-472">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-473">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-473">Requirements</span></span>

|<span data-ttu-id="a705f-474">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-474">Requirement</span></span>|<span data-ttu-id="a705f-475">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-476">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-476">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-477">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-477">1.0</span></span>|
|[<span data-ttu-id="a705f-478">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-478">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-479">ReadItem</span></span>|
|[<span data-ttu-id="a705f-480">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-480">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-481">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-481">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-482">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-482">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="a705f-483">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="a705f-483">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="a705f-484">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="a705f-484">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-485">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-485">Type:</span></span>

*   [<span data-ttu-id="a705f-486">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="a705f-486">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="a705f-487">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-487">Requirements</span></span>

|<span data-ttu-id="a705f-488">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-488">Requirement</span></span>|<span data-ttu-id="a705f-489">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-490">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-490">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-491">1.3</span><span class="sxs-lookup"><span data-stu-id="a705f-491">1.3</span></span>|
|[<span data-ttu-id="a705f-492">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-492">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-493">ReadItem</span></span>|
|[<span data-ttu-id="a705f-494">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-494">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-495">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-495">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a705f-496">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a705f-496">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a705f-497">Fournit l’accès à un événement les participants facultatifs.</span><span class="sxs-lookup"><span data-stu-id="a705f-497">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="a705f-498">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="a705f-498">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a705f-499">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-499">Read mode</span></span>

<span data-ttu-id="a705f-500">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="a705f-500">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a705f-501">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a705f-501">Compose mode</span></span>

<span data-ttu-id="a705f-502">Le `optionalAttendees` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs à une réunion.</span><span class="sxs-lookup"><span data-stu-id="a705f-502">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-503">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-503">Type:</span></span>

*   <span data-ttu-id="a705f-504">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a705f-504">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-505">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-505">Requirements</span></span>

|<span data-ttu-id="a705f-506">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-506">Requirement</span></span>|<span data-ttu-id="a705f-507">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-508">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-508">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-509">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-509">1.0</span></span>|
|[<span data-ttu-id="a705f-510">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-511">ReadItem</span></span>|
|[<span data-ttu-id="a705f-512">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-513">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-514">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-514">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="a705f-515">organisateur :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[organisateur](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="a705f-515">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="a705f-516">Obtient l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="a705f-516">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a705f-517">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-517">Read mode</span></span>

<span data-ttu-id="a705f-518">Le `organizer` propriété renvoie un objet [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="a705f-518">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a705f-519">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a705f-519">Compose mode</span></span>

<span data-ttu-id="a705f-520">Le `organizer` propriété retourne un objet [organisateur](/javascript/api/outlook/office.organizer) qui fournit une méthode pour obtenir la valeur de l’organisateur.</span><span class="sxs-lookup"><span data-stu-id="a705f-520">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-521">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-521">Type:</span></span>

*   <span data-ttu-id="a705f-522">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [organisateur](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="a705f-522">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-523">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-523">Requirements</span></span>

|<span data-ttu-id="a705f-524">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-524">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a705f-525">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-525">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-526">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-526">1.0</span></span>|<span data-ttu-id="a705f-527">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a705f-527">Preview</span></span>|
|[<span data-ttu-id="a705f-528">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-528">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-529">ReadItem</span></span>|<span data-ttu-id="a705f-530">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a705f-530">ReadWriteItem</span></span>|
|[<span data-ttu-id="a705f-531">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-531">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-532">Read</span><span class="sxs-lookup"><span data-stu-id="a705f-532">Read</span></span>|<span data-ttu-id="a705f-533">Composition</span><span class="sxs-lookup"><span data-stu-id="a705f-533">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-534">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-534">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="a705f-535">périodicité (nullable) :[périodicité](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="a705f-535">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="a705f-536">Obtient ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a705f-536">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="a705f-537">Obtient la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="a705f-537">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="a705f-538">Lecture et de composition modes pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a705f-538">Read and compose modes for appointment items.</span></span> <span data-ttu-id="a705f-539">Mode de lecture pour les éléments de la demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="a705f-539">Read mode for meeting request items.</span></span>

<span data-ttu-id="a705f-540">Le `recurrence` propriété renvoie un objet de [périodicité](/javascript/api/outlook/office.recurrence) pour les demandes de réunions ou de rendez-vous périodiques si un élément est une série ou à une instance d’une série.</span><span class="sxs-lookup"><span data-stu-id="a705f-540">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="a705f-541">`null`sont retournées pour les rendez-vous uniques et des demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="a705f-541">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="a705f-542">`undefined`sont retournées pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="a705f-542">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="a705f-543">Remarque : Les demandes de réunion ont une `itemClass` valeur IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="a705f-543">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="a705f-544">Remarque : Si l’objet de périodicité est `null`, cela indique que l’objet est un rendez-vous ou une demande de réunion d’un rendez-vous et ne fait pas partie d’une série.</span><span class="sxs-lookup"><span data-stu-id="a705f-544">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-545">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-545">Type:</span></span>

* [<span data-ttu-id="a705f-546">Périodicité</span><span class="sxs-lookup"><span data-stu-id="a705f-546">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="a705f-547">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-547">Requirement</span></span>|<span data-ttu-id="a705f-548">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-549">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-549">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-550">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a705f-550">Preview</span></span>|
|[<span data-ttu-id="a705f-551">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-551">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-552">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-552">ReadItem</span></span>|
|[<span data-ttu-id="a705f-553">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-554">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-554">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a705f-555">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a705f-555">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a705f-556">Fournit l’accès à des participants obligatoires d’un événement.</span><span class="sxs-lookup"><span data-stu-id="a705f-556">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="a705f-557">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="a705f-557">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a705f-558">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-558">Read mode</span></span>

<span data-ttu-id="a705f-559">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="a705f-559">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a705f-560">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a705f-560">Compose mode</span></span>

<span data-ttu-id="a705f-561">Le `requiredAttendees` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les participants obligatoires à une réunion.</span><span class="sxs-lookup"><span data-stu-id="a705f-561">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-562">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-562">Type:</span></span>

*   <span data-ttu-id="a705f-563">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a705f-563">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-564">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-564">Requirements</span></span>

|<span data-ttu-id="a705f-565">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-565">Requirement</span></span>|<span data-ttu-id="a705f-566">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-567">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-567">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-568">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-568">1.0</span></span>|
|[<span data-ttu-id="a705f-569">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-569">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-570">ReadItem</span></span>|
|[<span data-ttu-id="a705f-571">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-571">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-572">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-572">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-573">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-573">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="a705f-574">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a705f-574">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="a705f-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a705f-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="a705f-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="a705f-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-579">Le `recipientType` propriété de la `EmailAddressDetails` objet dans le `sender` propriété est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a705f-579">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-580">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-580">Type:</span></span>

*   [<span data-ttu-id="a705f-581">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a705f-581">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a705f-582">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-582">Requirements</span></span>

|<span data-ttu-id="a705f-583">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-583">Requirement</span></span>|<span data-ttu-id="a705f-584">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-585">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-585">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-586">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-586">1.0</span></span>|
|[<span data-ttu-id="a705f-587">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-588">ReadItem</span></span>|
|[<span data-ttu-id="a705f-589">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-590">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-590">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-591">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-591">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="a705f-592">seriesId (nullable) : chaîne</span><span class="sxs-lookup"><span data-stu-id="a705f-592">(nullable) seriesId :String</span></span>

<span data-ttu-id="a705f-593">Obtient l’id de la série appartenant à une instance.</span><span class="sxs-lookup"><span data-stu-id="a705f-593">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="a705f-594">Dans OWA et Outlook, le `seriesId` renvoie l’ID Exchange Web Services (EWS) de l’élément parent (série) auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="a705f-594">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="a705f-595">Toutefois, dans iOS et Android, le `seriesId` renvoie l’ID du reste de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="a705f-595">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-596">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="a705f-596">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a705f-597">Le `seriesId` propriété n’est pas identique à l’ID Outlook utilisées par l’API REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="a705f-597">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="a705f-598">Avant d’effectuer des appels d’API REST à l’aide de cette valeur, elle doit être convertie à l’aide de [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a705f-598">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a705f-599">Pour plus d’informations, voir [utiliser les API REST d’Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="a705f-599">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="a705f-600">Le `seriesId` propriété renvoie `null` pour les éléments qui n’ont pas d’éléments parents comme unique rendez-vous, les éléments de série, ou les demandes de réunion et renvoie `undefined` pour tous les éléments qui ne sont pas les demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="a705f-600">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-601">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-601">Type:</span></span>

* <span data-ttu-id="a705f-602">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a705f-602">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-603">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-603">Requirements</span></span>

|<span data-ttu-id="a705f-604">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-604">Requirement</span></span>|<span data-ttu-id="a705f-605">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-606">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-606">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-607">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a705f-607">Preview</span></span>|
|[<span data-ttu-id="a705f-608">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-609">ReadItem</span></span>|
|[<span data-ttu-id="a705f-610">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-611">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-611">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-612">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-612">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="a705f-613">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a705f-613">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="a705f-614">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a705f-614">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="a705f-p130">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="a705f-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a705f-617">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-617">Read mode</span></span>

<span data-ttu-id="a705f-618">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="a705f-618">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a705f-619">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a705f-619">Compose mode</span></span>

<span data-ttu-id="a705f-620">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="a705f-620">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="a705f-621">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="a705f-621">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-622">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-622">Type:</span></span>

*   <span data-ttu-id="a705f-623">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="a705f-623">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-624">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-624">Requirements</span></span>

|<span data-ttu-id="a705f-625">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-625">Requirement</span></span>|<span data-ttu-id="a705f-626">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-627">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-627">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-628">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-628">1.0</span></span>|
|[<span data-ttu-id="a705f-629">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-629">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-630">ReadItem</span></span>|
|[<span data-ttu-id="a705f-631">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-631">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-632">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-632">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-633">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-633">Example</span></span>

<span data-ttu-id="a705f-634">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="a705f-634">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="a705f-635">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a705f-635">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="a705f-636">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="a705f-636">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="a705f-637">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="a705f-637">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a705f-638">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-638">Read mode</span></span>

<span data-ttu-id="a705f-p131">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="a705f-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="a705f-641">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a705f-641">Compose mode</span></span>

<span data-ttu-id="a705f-642">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="a705f-642">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a705f-643">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-643">Type:</span></span>

*   <span data-ttu-id="a705f-644">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a705f-644">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-645">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-645">Requirements</span></span>

|<span data-ttu-id="a705f-646">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-646">Requirement</span></span>|<span data-ttu-id="a705f-647">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-648">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-648">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-649">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-649">1.0</span></span>|
|[<span data-ttu-id="a705f-650">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-650">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-651">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-651">ReadItem</span></span>|
|[<span data-ttu-id="a705f-652">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-652">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-653">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-653">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="a705f-654">à : Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a705f-654">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="a705f-655">Permet d’accéder aux destinataires de la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="a705f-655">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="a705f-656">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="a705f-656">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a705f-657">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-657">Read mode</span></span>

<span data-ttu-id="a705f-p133">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="a705f-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a705f-660">Mode composition</span><span class="sxs-lookup"><span data-stu-id="a705f-660">Compose mode</span></span>

<span data-ttu-id="a705f-661">Le `to` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="a705f-661">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a705f-662">Type :</span><span class="sxs-lookup"><span data-stu-id="a705f-662">Type:</span></span>

*   <span data-ttu-id="a705f-663">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a705f-663">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-664">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-664">Requirements</span></span>

|<span data-ttu-id="a705f-665">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-665">Requirement</span></span>|<span data-ttu-id="a705f-666">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-666">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-667">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-667">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-668">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-668">1.0</span></span>|
|[<span data-ttu-id="a705f-669">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-669">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-670">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-670">ReadItem</span></span>|
|[<span data-ttu-id="a705f-671">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-671">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-672">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-672">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-673">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-673">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="a705f-674">Méthodes</span><span class="sxs-lookup"><span data-stu-id="a705f-674">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="a705f-675">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a705f-675">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a705f-676">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="a705f-676">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a705f-677">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="a705f-677">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="a705f-678">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="a705f-678">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-679">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-679">Parameters:</span></span>
|<span data-ttu-id="a705f-680">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-680">Name</span></span>|<span data-ttu-id="a705f-681">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-681">Type</span></span>|<span data-ttu-id="a705f-682">Attributs</span><span class="sxs-lookup"><span data-stu-id="a705f-682">Attributes</span></span>|<span data-ttu-id="a705f-683">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-683">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="a705f-684">String</span><span class="sxs-lookup"><span data-stu-id="a705f-684">String</span></span>||<span data-ttu-id="a705f-p134">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="a705f-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a705f-687">String</span><span class="sxs-lookup"><span data-stu-id="a705f-687">String</span></span>||<span data-ttu-id="a705f-p135">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a705f-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a705f-690">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-690">Object</span></span>|<span data-ttu-id="a705f-691">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-691">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-692">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a705f-692">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a705f-693">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-693">Object</span></span>|<span data-ttu-id="a705f-694">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-694">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-695">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-695">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="a705f-696">Boolean</span><span class="sxs-lookup"><span data-stu-id="a705f-696">Boolean</span></span>|<span data-ttu-id="a705f-697">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-697">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-698">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a705f-698">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="a705f-699">fonction</span><span class="sxs-lookup"><span data-stu-id="a705f-699">function</span></span>|<span data-ttu-id="a705f-700">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-700">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-701">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a705f-701">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a705f-702">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a705f-702">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a705f-703">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="a705f-703">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a705f-704">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a705f-704">Errors</span></span>

|<span data-ttu-id="a705f-705">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a705f-705">Error code</span></span>|<span data-ttu-id="a705f-706">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-706">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="a705f-707">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="a705f-707">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="a705f-708">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="a705f-708">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a705f-709">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a705f-709">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-710">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-710">Requirements</span></span>

|<span data-ttu-id="a705f-711">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-711">Requirement</span></span>|<span data-ttu-id="a705f-712">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-712">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-713">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-713">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-714">1.1</span><span class="sxs-lookup"><span data-stu-id="a705f-714">1.1</span></span>|
|[<span data-ttu-id="a705f-715">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-715">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-716">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a705f-716">ReadWriteItem</span></span>|
|[<span data-ttu-id="a705f-717">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-717">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-718">Composition</span><span class="sxs-lookup"><span data-stu-id="a705f-718">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a705f-719">範例</span><span class="sxs-lookup"><span data-stu-id="a705f-719">Examples</span></span>

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

<span data-ttu-id="a705f-720">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="a705f-720">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="a705f-721">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [rappel])</span><span class="sxs-lookup"><span data-stu-id="a705f-721">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a705f-722">Ajoute un fichier à partir de l’en base64 codage à un message ou un rendez-vous en tant que pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="a705f-722">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a705f-723">Le `addFileAttachmentFromBase64Async` méthode télécharge le fichier à partir du codage en base64 et l’attache à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="a705f-723">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="a705f-724">Cette méthode renvoie l’identificateur de pièce jointe dans l’objet AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="a705f-724">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="a705f-725">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="a705f-725">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-726">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-726">Parameters:</span></span>
|<span data-ttu-id="a705f-727">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-727">Name</span></span>|<span data-ttu-id="a705f-728">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-728">Type</span></span>|<span data-ttu-id="a705f-729">Attributs</span><span class="sxs-lookup"><span data-stu-id="a705f-729">Attributes</span></span>|<span data-ttu-id="a705f-730">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-730">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="a705f-731">String</span><span class="sxs-lookup"><span data-stu-id="a705f-731">String</span></span>||<span data-ttu-id="a705f-732">Le contenu d’une image ou un fichier à ajouter à un courrier électronique ou un événement codage Base64.</span><span class="sxs-lookup"><span data-stu-id="a705f-732">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="a705f-733">String</span><span class="sxs-lookup"><span data-stu-id="a705f-733">String</span></span>||<span data-ttu-id="a705f-p137">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a705f-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a705f-736">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-736">Object</span></span>|<span data-ttu-id="a705f-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-737">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-738">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a705f-738">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a705f-739">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-739">Object</span></span>|<span data-ttu-id="a705f-740">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-740">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-741">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-741">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="a705f-742">Boolean</span><span class="sxs-lookup"><span data-stu-id="a705f-742">Boolean</span></span>|<span data-ttu-id="a705f-743">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-743">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-744">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a705f-744">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="a705f-745">fonction</span><span class="sxs-lookup"><span data-stu-id="a705f-745">function</span></span>|<span data-ttu-id="a705f-746">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-746">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-747">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a705f-747">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a705f-748">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a705f-748">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a705f-749">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="a705f-749">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a705f-750">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a705f-750">Errors</span></span>

|<span data-ttu-id="a705f-751">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a705f-751">Error code</span></span>|<span data-ttu-id="a705f-752">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-752">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="a705f-753">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="a705f-753">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="a705f-754">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="a705f-754">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a705f-755">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a705f-755">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-756">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-756">Requirements</span></span>

|<span data-ttu-id="a705f-757">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-757">Requirement</span></span>|<span data-ttu-id="a705f-758">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-758">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-759">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-759">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-760">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a705f-760">Preview</span></span>|
|[<span data-ttu-id="a705f-761">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-761">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-762">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a705f-762">ReadWriteItem</span></span>|
|[<span data-ttu-id="a705f-763">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-763">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-764">Composition</span><span class="sxs-lookup"><span data-stu-id="a705f-764">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a705f-765">範例</span><span class="sxs-lookup"><span data-stu-id="a705f-765">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a705f-766">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a705f-766">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a705f-767">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="a705f-767">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="a705f-768">Actuellement les types d’événements pris en charge sont `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, et`Office.EventType.RecurrencePatternChanged`</span><span class="sxs-lookup"><span data-stu-id="a705f-768">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrencePatternChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-769">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-769">Parameters:</span></span>

| <span data-ttu-id="a705f-770">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-770">Name</span></span> | <span data-ttu-id="a705f-771">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-771">Type</span></span> | <span data-ttu-id="a705f-772">Attributs</span><span class="sxs-lookup"><span data-stu-id="a705f-772">Attributes</span></span> | <span data-ttu-id="a705f-773">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-773">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a705f-774">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a705f-774">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a705f-775">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="a705f-775">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a705f-776">Fonction</span><span class="sxs-lookup"><span data-stu-id="a705f-776">Function</span></span> || <span data-ttu-id="a705f-p138">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="a705f-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a705f-780">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-780">Object</span></span> | <span data-ttu-id="a705f-781">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-781">&lt;optional&gt;</span></span> | <span data-ttu-id="a705f-782">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a705f-782">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a705f-783">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-783">Object</span></span> | <span data-ttu-id="a705f-784">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-784">&lt;optional&gt;</span></span> | <span data-ttu-id="a705f-785">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-785">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a705f-786">fonction</span><span class="sxs-lookup"><span data-stu-id="a705f-786">function</span></span>| <span data-ttu-id="a705f-787">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-787">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-788">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a705f-788">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-789">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-789">Requirements</span></span>

|<span data-ttu-id="a705f-790">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-790">Requirement</span></span>| <span data-ttu-id="a705f-791">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-791">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-792">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-792">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a705f-793">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a705f-793">Preview</span></span> |
|[<span data-ttu-id="a705f-794">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-794">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a705f-795">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-795">ReadItem</span></span> |
|[<span data-ttu-id="a705f-796">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-796">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a705f-797">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-797">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="a705f-798">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-798">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrencePatternChanged, loadNewItem, function (result) {
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="a705f-799">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a705f-799">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a705f-800">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a705f-800">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="a705f-p139">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="a705f-804">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="a705f-804">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="a705f-805">Si votre complément Office est en cours d’exécution dans Outlook Web App, le `addItemAttachmentAsync` méthode permettre joindre des éléments à des éléments autres que l’élément à modifier ; Toutefois, cela n’est pas pris en charge et n’est pas recommandée.</span><span class="sxs-lookup"><span data-stu-id="a705f-805">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-806">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-806">Parameters:</span></span>

|<span data-ttu-id="a705f-807">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-807">Name</span></span>|<span data-ttu-id="a705f-808">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-808">Type</span></span>|<span data-ttu-id="a705f-809">Attributs</span><span class="sxs-lookup"><span data-stu-id="a705f-809">Attributes</span></span>|<span data-ttu-id="a705f-810">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-810">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="a705f-811">String</span><span class="sxs-lookup"><span data-stu-id="a705f-811">String</span></span>||<span data-ttu-id="a705f-p140">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="a705f-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a705f-814">String</span><span class="sxs-lookup"><span data-stu-id="a705f-814">String</span></span>||<span data-ttu-id="a705f-p141">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a705f-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a705f-817">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-817">Object</span></span>|<span data-ttu-id="a705f-818">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-818">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-819">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a705f-819">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a705f-820">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-820">Object</span></span>|<span data-ttu-id="a705f-821">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-821">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-822">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-822">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a705f-823">fonction</span><span class="sxs-lookup"><span data-stu-id="a705f-823">function</span></span>|<span data-ttu-id="a705f-824">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-824">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-825">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a705f-825">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a705f-826">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a705f-826">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a705f-827">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="a705f-827">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a705f-828">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a705f-828">Errors</span></span>

|<span data-ttu-id="a705f-829">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a705f-829">Error code</span></span>|<span data-ttu-id="a705f-830">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-830">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a705f-831">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a705f-831">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-832">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-832">Requirements</span></span>

|<span data-ttu-id="a705f-833">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-833">Requirement</span></span>|<span data-ttu-id="a705f-834">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-835">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-835">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-836">1.1</span><span class="sxs-lookup"><span data-stu-id="a705f-836">1.1</span></span>|
|[<span data-ttu-id="a705f-837">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-838">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a705f-838">ReadWriteItem</span></span>|
|[<span data-ttu-id="a705f-839">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-840">Composition</span><span class="sxs-lookup"><span data-stu-id="a705f-840">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-841">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-841">Example</span></span>

<span data-ttu-id="a705f-842">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="a705f-842">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="a705f-843">close()</span><span class="sxs-lookup"><span data-stu-id="a705f-843">close()</span></span>

<span data-ttu-id="a705f-844">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="a705f-844">Closes the current item that is being composed.</span></span>

<span data-ttu-id="a705f-p142">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="a705f-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-847">Dans Outlook sur le web, si l’élément est un rendez-vous et il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, supprimer ou annuler même si aucune modification de l’élément depuis le dernier enregistrée.</span><span class="sxs-lookup"><span data-stu-id="a705f-847">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="a705f-848">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="a705f-848">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-849">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-849">Requirements</span></span>

|<span data-ttu-id="a705f-850">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-850">Requirement</span></span>|<span data-ttu-id="a705f-851">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-851">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-852">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-852">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-853">1.3</span><span class="sxs-lookup"><span data-stu-id="a705f-853">1.3</span></span>|
|[<span data-ttu-id="a705f-854">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-854">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-855">Restricted</span><span class="sxs-lookup"><span data-stu-id="a705f-855">Restricted</span></span>|
|[<span data-ttu-id="a705f-856">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-856">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-857">Composition</span><span class="sxs-lookup"><span data-stu-id="a705f-857">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="a705f-858">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a705f-858">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="a705f-859">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a705f-859">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-860">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a705f-860">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a705f-861">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="a705f-861">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a705f-862">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="a705f-862">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="a705f-p143">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="a705f-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-866">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-866">Parameters:</span></span>

|<span data-ttu-id="a705f-867">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-867">Name</span></span>|<span data-ttu-id="a705f-868">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-868">Type</span></span>|<span data-ttu-id="a705f-869">Attributs</span><span class="sxs-lookup"><span data-stu-id="a705f-869">Attributes</span></span>|<span data-ttu-id="a705f-870">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-870">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a705f-871">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a705f-871">String &#124; Object</span></span>||<span data-ttu-id="a705f-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="a705f-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a705f-874">**OU**</span><span class="sxs-lookup"><span data-stu-id="a705f-874">**OR**</span></span><br/><span data-ttu-id="a705f-p145">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="a705f-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a705f-877">String</span><span class="sxs-lookup"><span data-stu-id="a705f-877">String</span></span>|<span data-ttu-id="a705f-878">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-878">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="a705f-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a705f-881">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-881">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a705f-882">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-882">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-883">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="a705f-883">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a705f-884">String</span><span class="sxs-lookup"><span data-stu-id="a705f-884">String</span></span>||<span data-ttu-id="a705f-p147">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="a705f-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a705f-887">String</span><span class="sxs-lookup"><span data-stu-id="a705f-887">String</span></span>||<span data-ttu-id="a705f-888">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a705f-888">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a705f-889">String</span><span class="sxs-lookup"><span data-stu-id="a705f-889">String</span></span>||<span data-ttu-id="a705f-p148">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="a705f-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a705f-892">Boolean</span><span class="sxs-lookup"><span data-stu-id="a705f-892">Boolean</span></span>||<span data-ttu-id="a705f-p149">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a705f-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a705f-895">String</span><span class="sxs-lookup"><span data-stu-id="a705f-895">String</span></span>||<span data-ttu-id="a705f-p150">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="a705f-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a705f-899">function</span><span class="sxs-lookup"><span data-stu-id="a705f-899">function</span></span>|<span data-ttu-id="a705f-900">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-900">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-901">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a705f-901">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-902">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-902">Requirements</span></span>

|<span data-ttu-id="a705f-903">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-903">Requirement</span></span>|<span data-ttu-id="a705f-904">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-905">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-906">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-906">1.0</span></span>|
|[<span data-ttu-id="a705f-907">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-908">ReadItem</span></span>|
|[<span data-ttu-id="a705f-909">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-910">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-910">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a705f-911">Exemples</span><span class="sxs-lookup"><span data-stu-id="a705f-911">Examples</span></span>

<span data-ttu-id="a705f-912">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="a705f-912">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="a705f-913">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="a705f-913">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="a705f-914">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="a705f-914">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a705f-915">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="a705f-915">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a705f-916">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="a705f-916">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a705f-917">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-917">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="a705f-918">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a705f-918">displayReplyForm(formData)</span></span>

<span data-ttu-id="a705f-919">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a705f-919">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-920">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a705f-920">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a705f-921">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="a705f-921">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a705f-922">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="a705f-922">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="a705f-p151">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="a705f-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-926">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-926">Parameters:</span></span>

|<span data-ttu-id="a705f-927">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-927">Name</span></span>|<span data-ttu-id="a705f-928">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-928">Type</span></span>|<span data-ttu-id="a705f-929">Attributs</span><span class="sxs-lookup"><span data-stu-id="a705f-929">Attributes</span></span>|<span data-ttu-id="a705f-930">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-930">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a705f-931">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a705f-931">String &#124; Object</span></span>||<span data-ttu-id="a705f-p152">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="a705f-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a705f-934">**OU**</span><span class="sxs-lookup"><span data-stu-id="a705f-934">**OR**</span></span><br/><span data-ttu-id="a705f-p153">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="a705f-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a705f-937">String</span><span class="sxs-lookup"><span data-stu-id="a705f-937">String</span></span>|<span data-ttu-id="a705f-938">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-938">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="a705f-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a705f-941">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-941">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a705f-942">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-942">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-943">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="a705f-943">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a705f-944">String</span><span class="sxs-lookup"><span data-stu-id="a705f-944">String</span></span>||<span data-ttu-id="a705f-p155">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="a705f-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a705f-947">String</span><span class="sxs-lookup"><span data-stu-id="a705f-947">String</span></span>||<span data-ttu-id="a705f-948">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="a705f-948">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a705f-949">String</span><span class="sxs-lookup"><span data-stu-id="a705f-949">String</span></span>||<span data-ttu-id="a705f-p156">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="a705f-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a705f-952">Boolean</span><span class="sxs-lookup"><span data-stu-id="a705f-952">Boolean</span></span>||<span data-ttu-id="a705f-p157">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="a705f-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a705f-955">String</span><span class="sxs-lookup"><span data-stu-id="a705f-955">String</span></span>||<span data-ttu-id="a705f-p158">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="a705f-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a705f-959">function</span><span class="sxs-lookup"><span data-stu-id="a705f-959">function</span></span>|<span data-ttu-id="a705f-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-960">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-961">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a705f-961">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-962">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-962">Requirements</span></span>

|<span data-ttu-id="a705f-963">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-963">Requirement</span></span>|<span data-ttu-id="a705f-964">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-964">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-965">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-965">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-966">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-966">1.0</span></span>|
|[<span data-ttu-id="a705f-967">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-967">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-968">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-968">ReadItem</span></span>|
|[<span data-ttu-id="a705f-969">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-969">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-970">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-970">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a705f-971">Exemples</span><span class="sxs-lookup"><span data-stu-id="a705f-971">Examples</span></span>

<span data-ttu-id="a705f-972">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="a705f-972">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="a705f-973">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="a705f-973">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="a705f-974">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="a705f-974">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a705f-975">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="a705f-975">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a705f-976">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="a705f-976">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a705f-977">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-977">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="a705f-978">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a705f-978">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="a705f-979">Obtient les entités trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a705f-979">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-980">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a705f-980">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-981">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-981">Requirements</span></span>

|<span data-ttu-id="a705f-982">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-982">Requirement</span></span>|<span data-ttu-id="a705f-983">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-984">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-984">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-985">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-985">1.0</span></span>|
|[<span data-ttu-id="a705f-986">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-986">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-987">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-987">ReadItem</span></span>|
|[<span data-ttu-id="a705f-988">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-988">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-989">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-989">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a705f-990">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a705f-990">Returns:</span></span>

<span data-ttu-id="a705f-991">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a705f-991">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a705f-992">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-992">Example</span></span>

<span data-ttu-id="a705f-993">L’exemple suivant accède aux entités de contacts dans le corps de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="a705f-993">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="a705f-994">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a705f-994">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a705f-995">Obtient un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a705f-995">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-996">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a705f-996">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-997">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-997">Parameters:</span></span>

|<span data-ttu-id="a705f-998">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-998">Name</span></span>|<span data-ttu-id="a705f-999">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-999">Type</span></span>|<span data-ttu-id="a705f-1000">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-1000">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="a705f-1001">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="a705f-1001">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="a705f-1002">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="a705f-1002">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-1003">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-1003">Requirements</span></span>

|<span data-ttu-id="a705f-1004">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-1004">Requirement</span></span>|<span data-ttu-id="a705f-1005">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-1006">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-1006">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-1007">1.0</span></span>|
|[<span data-ttu-id="a705f-1008">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-1008">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-1009">Restricted</span><span class="sxs-lookup"><span data-stu-id="a705f-1009">Restricted</span></span>|
|[<span data-ttu-id="a705f-1010">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-1010">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-1011">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a705f-1012">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a705f-1012">Returns:</span></span>

<span data-ttu-id="a705f-1013">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="a705f-1013">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="a705f-1014">Si aucune entité du type spécifié n’est présentes dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="a705f-1014">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="a705f-1015">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="a705f-1015">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="a705f-1016">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="a705f-1016">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="a705f-1017">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="a705f-1017">Value of `entityType`</span></span>|<span data-ttu-id="a705f-1018">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="a705f-1018">Type of objects in returned array</span></span>|<span data-ttu-id="a705f-1019">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="a705f-1019">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="a705f-1020">String</span><span class="sxs-lookup"><span data-stu-id="a705f-1020">String</span></span>|<span data-ttu-id="a705f-1021">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a705f-1021">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="a705f-1022">Contact</span><span class="sxs-lookup"><span data-stu-id="a705f-1022">Contact</span></span>|<span data-ttu-id="a705f-1023">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a705f-1023">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="a705f-1024">String</span><span class="sxs-lookup"><span data-stu-id="a705f-1024">String</span></span>|<span data-ttu-id="a705f-1025">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a705f-1025">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="a705f-1026">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="a705f-1026">MeetingSuggestion</span></span>|<span data-ttu-id="a705f-1027">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a705f-1027">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="a705f-1028">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="a705f-1028">PhoneNumber</span></span>|<span data-ttu-id="a705f-1029">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a705f-1029">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="a705f-1030">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="a705f-1030">TaskSuggestion</span></span>|<span data-ttu-id="a705f-1031">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a705f-1031">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="a705f-1032">String</span><span class="sxs-lookup"><span data-stu-id="a705f-1032">String</span></span>|<span data-ttu-id="a705f-1033">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a705f-1033">**Restricted**</span></span>|

<span data-ttu-id="a705f-1034">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a705f-1034">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="a705f-1035">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-1035">Example</span></span>

<span data-ttu-id="a705f-1036">L’exemple suivant montre comment accéder à un tableau de chaînes qui représentent les adresses postales dans le corps de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="a705f-1036">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="a705f-1037">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a705f-1037">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a705f-1038">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a705f-1038">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-1039">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a705f-1039">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a705f-1040">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="a705f-1040">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-1041">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-1041">Parameters:</span></span>

|<span data-ttu-id="a705f-1042">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-1042">Name</span></span>|<span data-ttu-id="a705f-1043">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-1043">Type</span></span>|<span data-ttu-id="a705f-1044">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-1044">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a705f-1045">String</span><span class="sxs-lookup"><span data-stu-id="a705f-1045">String</span></span>|<span data-ttu-id="a705f-1046">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="a705f-1046">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-1047">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-1047">Requirements</span></span>

|<span data-ttu-id="a705f-1048">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-1048">Requirement</span></span>|<span data-ttu-id="a705f-1049">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-1049">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-1050">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-1050">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-1051">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-1051">1.0</span></span>|
|[<span data-ttu-id="a705f-1052">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-1052">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-1053">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-1053">ReadItem</span></span>|
|[<span data-ttu-id="a705f-1054">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-1054">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-1055">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-1055">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a705f-1056">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a705f-1056">Returns:</span></span>

<span data-ttu-id="a705f-p160">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="a705f-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="a705f-1059">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a705f-1059">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="a705f-1060">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a705f-1060">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="a705f-1061">Récupère les données d’initialisation transmises quand le complément est [activé par un message actionnable](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="a705f-1061">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-1062">Cette méthode est uniquement pris en charge par Outlook 2016 pour Windows (Click-to-Run versions supérieures à 16.0.8413.1000) et Outlook sur le web pour Office 365.</span><span class="sxs-lookup"><span data-stu-id="a705f-1062">This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-1063">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-1063">Parameters:</span></span>
|<span data-ttu-id="a705f-1064">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-1064">Name</span></span>|<span data-ttu-id="a705f-1065">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-1065">Type</span></span>|<span data-ttu-id="a705f-1066">Attributs</span><span class="sxs-lookup"><span data-stu-id="a705f-1066">Attributes</span></span>|<span data-ttu-id="a705f-1067">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-1067">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a705f-1068">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-1068">Object</span></span>|<span data-ttu-id="a705f-1069">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1070">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a705f-1070">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a705f-1071">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-1071">Object</span></span>|<span data-ttu-id="a705f-1072">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1072">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1073">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-1073">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a705f-1074">fonction</span><span class="sxs-lookup"><span data-stu-id="a705f-1074">function</span></span>|<span data-ttu-id="a705f-1075">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1075">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1076">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a705f-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a705f-1077">En cas de réussite, les données d’initialisation sont fournies dans le `asyncResult.value` propriété sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="a705f-1077">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="a705f-1078">S’il n’existe aucun contexte d’initialisation, l’objet `asyncResult` contient un objet `Error` dont la propriété `code` est définie sur `9020` et la propriété `name` sur `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="a705f-1078">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-1079">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-1079">Requirements</span></span>

|<span data-ttu-id="a705f-1080">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-1080">Requirement</span></span>|<span data-ttu-id="a705f-1081">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-1081">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-1082">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-1082">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-1083">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a705f-1083">Preview</span></span>|
|[<span data-ttu-id="a705f-1084">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-1084">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-1085">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-1085">ReadItem</span></span>|
|[<span data-ttu-id="a705f-1086">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-1086">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-1087">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-1087">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-1088">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-1088">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="a705f-1089">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a705f-1089">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="a705f-1090">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a705f-1090">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-1091">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a705f-1091">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a705f-p161">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="a705f-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a705f-1095">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="a705f-1095">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a705f-1096">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a705f-1096">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a705f-p162">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a705f-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-1100">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-1100">Requirements</span></span>

|<span data-ttu-id="a705f-1101">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-1101">Requirement</span></span>|<span data-ttu-id="a705f-1102">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-1102">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-1103">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-1103">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-1104">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-1104">1.0</span></span>|
|[<span data-ttu-id="a705f-1105">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-1105">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-1106">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-1106">ReadItem</span></span>|
|[<span data-ttu-id="a705f-1107">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-1107">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-1108">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-1108">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a705f-1109">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a705f-1109">Returns:</span></span>

<span data-ttu-id="a705f-p163">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="a705f-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="a705f-1112">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="a705f-1112">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a705f-1113">Object</span><span class="sxs-lookup"><span data-stu-id="a705f-1113">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a705f-1114">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-1114">Example</span></span>

<span data-ttu-id="a705f-1115">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="a705f-1115">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="a705f-1116">getregexmatchesbyname (Name) → (nullable) {tableau. < chaîne >}</span><span class="sxs-lookup"><span data-stu-id="a705f-1116">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="a705f-1117">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a705f-1117">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-1118">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a705f-1118">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a705f-1119">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="a705f-1119">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="a705f-p164">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="a705f-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-1122">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-1122">Parameters:</span></span>

|<span data-ttu-id="a705f-1123">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-1123">Name</span></span>|<span data-ttu-id="a705f-1124">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-1124">Type</span></span>|<span data-ttu-id="a705f-1125">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-1125">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a705f-1126">String</span><span class="sxs-lookup"><span data-stu-id="a705f-1126">String</span></span>|<span data-ttu-id="a705f-1127">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="a705f-1127">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-1128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-1128">Requirements</span></span>

|<span data-ttu-id="a705f-1129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-1129">Requirement</span></span>|<span data-ttu-id="a705f-1130">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-1130">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-1131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-1131">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-1132">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-1132">1.0</span></span>|
|[<span data-ttu-id="a705f-1133">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-1133">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-1134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-1134">ReadItem</span></span>|
|[<span data-ttu-id="a705f-1135">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-1135">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-1136">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-1136">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a705f-1137">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a705f-1137">Returns:</span></span>

<span data-ttu-id="a705f-1138">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="a705f-1138">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="a705f-1139">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="a705f-1139">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a705f-1140">< Chaîne > tableau.</span><span class="sxs-lookup"><span data-stu-id="a705f-1140">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a705f-1141">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-1141">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="a705f-1142">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="a705f-1142">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="a705f-1143">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="a705f-1143">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="a705f-p165">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="a705f-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-1146">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-1146">Parameters:</span></span>

|<span data-ttu-id="a705f-1147">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-1147">Name</span></span>|<span data-ttu-id="a705f-1148">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-1148">Type</span></span>|<span data-ttu-id="a705f-1149">Attributs</span><span class="sxs-lookup"><span data-stu-id="a705f-1149">Attributes</span></span>|<span data-ttu-id="a705f-1150">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-1150">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="a705f-1151">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a705f-1151">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="a705f-p166">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="a705f-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="a705f-1155">Object</span><span class="sxs-lookup"><span data-stu-id="a705f-1155">Object</span></span>|<span data-ttu-id="a705f-1156">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1156">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1157">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a705f-1157">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a705f-1158">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-1158">Object</span></span>|<span data-ttu-id="a705f-1159">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1160">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-1160">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a705f-1161">fonction</span><span class="sxs-lookup"><span data-stu-id="a705f-1161">function</span></span>||<span data-ttu-id="a705f-1162">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a705f-1162">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a705f-1163">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="a705f-1163">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="a705f-1164">Pour accéder à la propriété source provenant de la sélection, appelez `asyncResult.value.sourceProperty`, qui sera soit `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="a705f-1164">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-1165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-1165">Requirements</span></span>

|<span data-ttu-id="a705f-1166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-1166">Requirement</span></span>|<span data-ttu-id="a705f-1167">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-1167">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-1168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-1168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-1169">1.2</span><span class="sxs-lookup"><span data-stu-id="a705f-1169">1.2</span></span>|
|[<span data-ttu-id="a705f-1170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-1170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-1171">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a705f-1171">ReadWriteItem</span></span>|
|[<span data-ttu-id="a705f-1172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-1172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-1173">Composition</span><span class="sxs-lookup"><span data-stu-id="a705f-1173">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a705f-1174">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a705f-1174">Returns:</span></span>

<span data-ttu-id="a705f-1175">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="a705f-1175">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="a705f-1176">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="a705f-1176">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a705f-1177">String</span><span class="sxs-lookup"><span data-stu-id="a705f-1177">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a705f-1178">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-1178">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="a705f-1179">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a705f-1179">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="a705f-p168">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a705f-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-1182">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a705f-1182">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-1183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-1183">Requirements</span></span>

|<span data-ttu-id="a705f-1184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-1184">Requirement</span></span>|<span data-ttu-id="a705f-1185">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-1185">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-1186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-1186">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-1187">1.6</span><span class="sxs-lookup"><span data-stu-id="a705f-1187">1.6</span></span>|
|[<span data-ttu-id="a705f-1188">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-1188">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-1189">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-1189">ReadItem</span></span>|
|[<span data-ttu-id="a705f-1190">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-1190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-1191">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-1191">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a705f-1192">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a705f-1192">Returns:</span></span>

<span data-ttu-id="a705f-1193">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a705f-1193">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a705f-1194">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-1194">Example</span></span>

<span data-ttu-id="a705f-1195">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a705f-1195">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="a705f-1196">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a705f-1196">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="a705f-p169">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a705f-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-1199">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a705f-1199">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a705f-p170">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="a705f-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a705f-1203">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="a705f-1203">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a705f-1204">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a705f-1204">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a705f-p171">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a705f-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a705f-1208">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-1208">Requirements</span></span>

|<span data-ttu-id="a705f-1209">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-1209">Requirement</span></span>|<span data-ttu-id="a705f-1210">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-1210">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-1211">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-1211">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-1212">1.6</span><span class="sxs-lookup"><span data-stu-id="a705f-1212">1.6</span></span>|
|[<span data-ttu-id="a705f-1213">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-1213">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-1214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-1214">ReadItem</span></span>|
|[<span data-ttu-id="a705f-1215">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-1215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-1216">Lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-1216">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a705f-1217">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a705f-1217">Returns:</span></span>

<span data-ttu-id="a705f-p172">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="a705f-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="a705f-1220">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-1220">Example</span></span>

<span data-ttu-id="a705f-1221">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="a705f-1221">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="a705f-1222">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a705f-1222">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="a705f-1223">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="a705f-1223">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="a705f-p173">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="a705f-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-1227">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-1227">Parameters:</span></span>

|<span data-ttu-id="a705f-1228">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-1228">Name</span></span>|<span data-ttu-id="a705f-1229">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-1229">Type</span></span>|<span data-ttu-id="a705f-1230">Attributs</span><span class="sxs-lookup"><span data-stu-id="a705f-1230">Attributes</span></span>|<span data-ttu-id="a705f-1231">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-1231">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="a705f-1232">function</span><span class="sxs-lookup"><span data-stu-id="a705f-1232">function</span></span>||<span data-ttu-id="a705f-1233">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a705f-1233">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a705f-1234">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a705f-1234">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a705f-1235">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées de l’élément et enregistrer les modifications à la propriété personnalisée est redéfinie sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="a705f-1235">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="a705f-1236">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-1236">Object</span></span>|<span data-ttu-id="a705f-1237">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1237">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1238">Les développeurs peuvent fournir n’importe quel objet qu’ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-1238">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="a705f-1239">Cet objet est accessible par le `asyncResult.asyncContext` propriété dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-1239">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-1240">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-1240">Requirements</span></span>

|<span data-ttu-id="a705f-1241">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-1241">Requirement</span></span>|<span data-ttu-id="a705f-1242">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-1243">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-1243">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-1244">1.0</span><span class="sxs-lookup"><span data-stu-id="a705f-1244">1.0</span></span>|
|[<span data-ttu-id="a705f-1245">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-1246">ReadItem</span></span>|
|[<span data-ttu-id="a705f-1247">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-1248">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-1249">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-1249">Example</span></span>

<span data-ttu-id="a705f-p176">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="a705f-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="a705f-1253">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a705f-1253">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="a705f-1254">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a705f-1254">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="a705f-p177">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="a705f-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-1259">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-1259">Parameters:</span></span>

|<span data-ttu-id="a705f-1260">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-1260">Name</span></span>|<span data-ttu-id="a705f-1261">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-1261">Type</span></span>|<span data-ttu-id="a705f-1262">Attributs</span><span class="sxs-lookup"><span data-stu-id="a705f-1262">Attributes</span></span>|<span data-ttu-id="a705f-1263">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-1263">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="a705f-1264">String</span><span class="sxs-lookup"><span data-stu-id="a705f-1264">String</span></span>||<span data-ttu-id="a705f-p178">Identificateur de la pièce jointe à supprimer. La longueur maximale de la chaîne est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="a705f-p178">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="a705f-1267">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-1267">Object</span></span>|<span data-ttu-id="a705f-1268">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1269">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a705f-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a705f-1270">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-1270">Object</span></span>|<span data-ttu-id="a705f-1271">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1272">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a705f-1273">fonction</span><span class="sxs-lookup"><span data-stu-id="a705f-1273">function</span></span>|<span data-ttu-id="a705f-1274">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1274">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1275">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a705f-1275">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a705f-1276">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="a705f-1276">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a705f-1277">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a705f-1277">Errors</span></span>

|<span data-ttu-id="a705f-1278">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a705f-1278">Error code</span></span>|<span data-ttu-id="a705f-1279">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-1279">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="a705f-1280">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="a705f-1280">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-1281">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-1281">Requirements</span></span>

|<span data-ttu-id="a705f-1282">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-1282">Requirement</span></span>|<span data-ttu-id="a705f-1283">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-1283">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-1284">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-1284">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-1285">1.1</span><span class="sxs-lookup"><span data-stu-id="a705f-1285">1.1</span></span>|
|[<span data-ttu-id="a705f-1286">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-1286">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-1287">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a705f-1287">ReadWriteItem</span></span>|
|[<span data-ttu-id="a705f-1288">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-1288">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-1289">Composition</span><span class="sxs-lookup"><span data-stu-id="a705f-1289">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-1290">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-1290">Example</span></span>

<span data-ttu-id="a705f-1291">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="a705f-1291">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a705f-1292">removeHandlerAsync (eventType, gestionnaire, [options], [rappel])</span><span class="sxs-lookup"><span data-stu-id="a705f-1292">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a705f-1293">Supprime un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="a705f-1293">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="a705f-1294">Actuellement les types d’événements pris en charge sont `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, et`Office.EventType.RecurrencePatternChanged`</span><span class="sxs-lookup"><span data-stu-id="a705f-1294">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrencePatternChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-1295">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-1295">Parameters:</span></span>

| <span data-ttu-id="a705f-1296">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-1296">Name</span></span> | <span data-ttu-id="a705f-1297">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-1297">Type</span></span> | <span data-ttu-id="a705f-1298">Attributs</span><span class="sxs-lookup"><span data-stu-id="a705f-1298">Attributes</span></span> | <span data-ttu-id="a705f-1299">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-1299">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a705f-1300">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a705f-1300">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a705f-1301">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="a705f-1301">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a705f-1302">Fonction</span><span class="sxs-lookup"><span data-stu-id="a705f-1302">Function</span></span> || <span data-ttu-id="a705f-p179">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="a705f-p179">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a705f-1306">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-1306">Object</span></span> | <span data-ttu-id="a705f-1307">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1307">&lt;optional&gt;</span></span> | <span data-ttu-id="a705f-1308">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a705f-1308">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a705f-1309">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-1309">Object</span></span> | <span data-ttu-id="a705f-1310">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1310">&lt;optional&gt;</span></span> | <span data-ttu-id="a705f-1311">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-1311">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a705f-1312">fonction</span><span class="sxs-lookup"><span data-stu-id="a705f-1312">function</span></span>| <span data-ttu-id="a705f-1313">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1313">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1314">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a705f-1314">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-1315">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-1315">Requirements</span></span>

|<span data-ttu-id="a705f-1316">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-1316">Requirement</span></span>| <span data-ttu-id="a705f-1317">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-1317">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-1318">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-1318">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a705f-1319">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a705f-1319">Preview</span></span> |
|[<span data-ttu-id="a705f-1320">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-1320">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a705f-1321">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a705f-1321">ReadItem</span></span> |
|[<span data-ttu-id="a705f-1322">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-1322">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a705f-1323">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a705f-1323">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="a705f-1324">Exemple</span><span class="sxs-lookup"><span data-stu-id="a705f-1324">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrencePatternChanged, loadNewItem, function (result) {
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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="a705f-1325">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="a705f-1325">saveAsync([options], callback)</span></span>

<span data-ttu-id="a705f-1326">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a705f-1326">Asynchronously saves an item.</span></span>

<span data-ttu-id="a705f-p180">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="a705f-p180">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-1330">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` pour utiliser avec EWS ou l’API REST, gardez à l’esprit que quand Outlook est en mode mis en cache, il peut prendre un certain temps avant que l’élément est réellement synchronisé avec le serveur.</span><span class="sxs-lookup"><span data-stu-id="a705f-1330">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="a705f-1331">Jusqu'à ce que l’élément est synchronisé, à l’aide de la `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="a705f-1331">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="a705f-p182">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="a705f-p182">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="a705f-1335">Les clients suivants ont un comportement différent pour `saveAsync` de rendez-vous dans le mode de composition :</span><span class="sxs-lookup"><span data-stu-id="a705f-1335">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="a705f-1336">Mac Outlook ne gère pas `saveAsync` d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="a705f-1336">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="a705f-1337">Appel de `saveAsync` d’une réunion dans Outlook Mac renverra une erreur.</span><span class="sxs-lookup"><span data-stu-id="a705f-1337">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="a705f-1338">Outlook sur le web toujours envoie une invitation ou mettre à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="a705f-1338">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-1339">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-1339">Parameters:</span></span>

|<span data-ttu-id="a705f-1340">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-1340">Name</span></span>|<span data-ttu-id="a705f-1341">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-1341">Type</span></span>|<span data-ttu-id="a705f-1342">Attributs</span><span class="sxs-lookup"><span data-stu-id="a705f-1342">Attributes</span></span>|<span data-ttu-id="a705f-1343">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-1343">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a705f-1344">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-1344">Object</span></span>|<span data-ttu-id="a705f-1345">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1345">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1346">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a705f-1346">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a705f-1347">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-1347">Object</span></span>|<span data-ttu-id="a705f-1348">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1348">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1349">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-1349">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a705f-1350">fonction</span><span class="sxs-lookup"><span data-stu-id="a705f-1350">function</span></span>||<span data-ttu-id="a705f-1351">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a705f-1351">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a705f-1352">En cas de réussite, l’identificateur d’élément est fournie dans le `asyncResult.value` propriété.</span><span class="sxs-lookup"><span data-stu-id="a705f-1352">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-1353">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-1353">Requirements</span></span>

|<span data-ttu-id="a705f-1354">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-1354">Requirement</span></span>|<span data-ttu-id="a705f-1355">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-1355">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-1356">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-1356">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-1357">1.3</span><span class="sxs-lookup"><span data-stu-id="a705f-1357">1.3</span></span>|
|[<span data-ttu-id="a705f-1358">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-1358">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-1359">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a705f-1359">ReadWriteItem</span></span>|
|[<span data-ttu-id="a705f-1360">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-1360">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-1361">Composition</span><span class="sxs-lookup"><span data-stu-id="a705f-1361">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a705f-1362">範例</span><span class="sxs-lookup"><span data-stu-id="a705f-1362">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="a705f-p184">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="a705f-p184">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="a705f-1365">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="a705f-1365">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="a705f-1366">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a705f-1366">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="a705f-p185">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="a705f-p185">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a705f-1370">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="a705f-1370">Parameters:</span></span>

|<span data-ttu-id="a705f-1371">Nom</span><span class="sxs-lookup"><span data-stu-id="a705f-1371">Name</span></span>|<span data-ttu-id="a705f-1372">Type</span><span class="sxs-lookup"><span data-stu-id="a705f-1372">Type</span></span>|<span data-ttu-id="a705f-1373">Attributs</span><span class="sxs-lookup"><span data-stu-id="a705f-1373">Attributes</span></span>|<span data-ttu-id="a705f-1374">Description</span><span class="sxs-lookup"><span data-stu-id="a705f-1374">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="a705f-1375">String</span><span class="sxs-lookup"><span data-stu-id="a705f-1375">String</span></span>||<span data-ttu-id="a705f-p186">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="a705f-p186">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="a705f-1379">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-1379">Object</span></span>|<span data-ttu-id="a705f-1380">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1380">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1381">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="a705f-1381">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a705f-1382">Objet</span><span class="sxs-lookup"><span data-stu-id="a705f-1382">Object</span></span>|<span data-ttu-id="a705f-1383">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1383">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-1384">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="a705f-1384">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="a705f-1385">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a705f-1385">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="a705f-1386">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a705f-1386">&lt;optional&gt;</span></span>|<span data-ttu-id="a705f-p187">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="a705f-p187">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="a705f-p188">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="a705f-p188">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="a705f-1391">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="a705f-1391">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="a705f-1392">fonction</span><span class="sxs-lookup"><span data-stu-id="a705f-1392">function</span></span>||<span data-ttu-id="a705f-1393">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a705f-1393">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a705f-1394">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a705f-1394">Requirements</span></span>

|<span data-ttu-id="a705f-1395">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a705f-1395">Requirement</span></span>|<span data-ttu-id="a705f-1396">Valeur</span><span class="sxs-lookup"><span data-stu-id="a705f-1396">Value</span></span>|
|---|---|
|[<span data-ttu-id="a705f-1397">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a705f-1397">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a705f-1398">1.2</span><span class="sxs-lookup"><span data-stu-id="a705f-1398">1.2</span></span>|
|[<span data-ttu-id="a705f-1399">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a705f-1399">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a705f-1400">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a705f-1400">ReadWriteItem</span></span>|
|[<span data-ttu-id="a705f-1401">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a705f-1401">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="a705f-1402">Composition</span><span class="sxs-lookup"><span data-stu-id="a705f-1402">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a705f-1403">Exemples</span><span class="sxs-lookup"><span data-stu-id="a705f-1403">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```