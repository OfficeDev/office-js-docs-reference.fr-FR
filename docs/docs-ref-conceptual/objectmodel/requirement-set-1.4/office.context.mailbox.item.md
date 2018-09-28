
# <a name="item"></a><span data-ttu-id="4472b-101">élément</span><span class="sxs-lookup"><span data-stu-id="4472b-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="4472b-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="4472b-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="4472b-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="4472b-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-105">Requirements</span></span>

|<span data-ttu-id="4472b-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-106">Requirement</span></span>| <span data-ttu-id="4472b-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-109">1.0</span></span>|
|[<span data-ttu-id="4472b-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-111">Restricted</span><span class="sxs-lookup"><span data-stu-id="4472b-111">Restricted</span></span>|
|[<span data-ttu-id="4472b-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="4472b-114">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-114">Example</span></span>

<span data-ttu-id="4472b-115">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="4472b-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="4472b-116">Membres</span><span class="sxs-lookup"><span data-stu-id="4472b-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="4472b-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4472b-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="4472b-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4472b-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-120">Certains types de fichiers bloqués par Outlook en raison de problèmes de sécurité potentiels et ne sont donc pas retournés.</span><span class="sxs-lookup"><span data-stu-id="4472b-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="4472b-121">Pour plus d’informations, voir les [pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="4472b-121">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-122">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-122">Type:</span></span>

*   <span data-ttu-id="4472b-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4472b-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-124">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-124">Requirements</span></span>

|<span data-ttu-id="4472b-125">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-125">Requirement</span></span>| <span data-ttu-id="4472b-126">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-127">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-127">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-128">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-128">1.0</span></span>|
|[<span data-ttu-id="4472b-129">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-130">ReadItem</span></span>|
|[<span data-ttu-id="4472b-131">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-133">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-133">Example</span></span>

<span data-ttu-id="4472b-134">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4472b-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="4472b-135">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4472b-135">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="4472b-136">Obtient un objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour celle-ci Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="4472b-136">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="4472b-137">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="4472b-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-138">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-138">Type:</span></span>

*   [<span data-ttu-id="4472b-139">Destinataires</span><span class="sxs-lookup"><span data-stu-id="4472b-139">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="4472b-140">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-140">Requirements</span></span>

|<span data-ttu-id="4472b-141">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-141">Requirement</span></span>| <span data-ttu-id="4472b-142">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-143">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-143">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-144">1.1</span><span class="sxs-lookup"><span data-stu-id="4472b-144">1.1</span></span>|
|[<span data-ttu-id="4472b-145">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-146">ReadItem</span></span>|
|[<span data-ttu-id="4472b-147">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-148">Composition</span><span class="sxs-lookup"><span data-stu-id="4472b-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-149">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-149">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="4472b-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="4472b-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="4472b-151">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4472b-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-152">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-152">Type:</span></span>

*   [<span data-ttu-id="4472b-153">Corps</span><span class="sxs-lookup"><span data-stu-id="4472b-153">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="4472b-154">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-154">Requirements</span></span>

|<span data-ttu-id="4472b-155">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-155">Requirement</span></span>| <span data-ttu-id="4472b-156">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-157">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-157">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-158">1.1</span><span class="sxs-lookup"><span data-stu-id="4472b-158">1.1</span></span>|
|[<span data-ttu-id="4472b-159">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-160">ReadItem</span></span>|
|[<span data-ttu-id="4472b-161">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-162">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="4472b-163">cc : tableau. <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4472b-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="4472b-164">Permet d’accéder aux destinataires Cc (copie carbone) d’un message.</span><span class="sxs-lookup"><span data-stu-id="4472b-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="4472b-165">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="4472b-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4472b-166">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-166">Read mode</span></span>

<span data-ttu-id="4472b-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="4472b-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4472b-169">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4472b-169">Compose mode</span></span>

<span data-ttu-id="4472b-170">Le `cc` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="4472b-170">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-171">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-171">Type:</span></span>

*   <span data-ttu-id="4472b-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4472b-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-173">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-173">Requirements</span></span>

|<span data-ttu-id="4472b-174">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-174">Requirement</span></span>| <span data-ttu-id="4472b-175">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-176">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-176">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-177">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-177">1.0</span></span>|
|[<span data-ttu-id="4472b-178">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-179">ReadItem</span></span>|
|[<span data-ttu-id="4472b-180">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-181">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-182">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-182">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="4472b-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="4472b-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="4472b-184">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="4472b-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="4472b-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="4472b-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="4472b-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="4472b-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-189">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-189">Type:</span></span>

*   <span data-ttu-id="4472b-190">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4472b-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-191">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-191">Requirements</span></span>

|<span data-ttu-id="4472b-192">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-192">Requirement</span></span>| <span data-ttu-id="4472b-193">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-194">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-195">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-195">1.0</span></span>|
|[<span data-ttu-id="4472b-196">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-197">ReadItem</span></span>|
|[<span data-ttu-id="4472b-198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-199">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="4472b-200">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="4472b-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="4472b-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4472b-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-203">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-203">Type:</span></span>

*   <span data-ttu-id="4472b-204">Date</span><span class="sxs-lookup"><span data-stu-id="4472b-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-205">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-205">Requirements</span></span>

|<span data-ttu-id="4472b-206">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-206">Requirement</span></span>| <span data-ttu-id="4472b-207">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-208">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-208">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-209">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-209">1.0</span></span>|
|[<span data-ttu-id="4472b-210">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-211">ReadItem</span></span>|
|[<span data-ttu-id="4472b-212">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-213">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-214">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-214">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="4472b-215">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="4472b-215">dateTimeModified :Date</span></span>

<span data-ttu-id="4472b-p110">Obtient la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4472b-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-218">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4472b-218">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-219">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-219">Type:</span></span>

*   <span data-ttu-id="4472b-220">Date</span><span class="sxs-lookup"><span data-stu-id="4472b-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-221">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-221">Requirements</span></span>

|<span data-ttu-id="4472b-222">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-222">Requirement</span></span>| <span data-ttu-id="4472b-223">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-224">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-225">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-225">1.0</span></span>|
|[<span data-ttu-id="4472b-226">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-227">ReadItem</span></span>|
|[<span data-ttu-id="4472b-228">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-229">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-230">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-230">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="4472b-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="4472b-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="4472b-232">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4472b-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="4472b-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="4472b-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4472b-235">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-235">Read mode</span></span>

<span data-ttu-id="4472b-236">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="4472b-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4472b-237">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4472b-237">Compose mode</span></span>

<span data-ttu-id="4472b-238">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4472b-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="4472b-239">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="4472b-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-240">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-240">Type:</span></span>

*   <span data-ttu-id="4472b-241">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="4472b-241">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-242">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-242">Requirements</span></span>

|<span data-ttu-id="4472b-243">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-243">Requirement</span></span>| <span data-ttu-id="4472b-244">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-245">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-245">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-246">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-246">1.0</span></span>|
|[<span data-ttu-id="4472b-247">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-248">ReadItem</span></span>|
|[<span data-ttu-id="4472b-249">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-250">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-251">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-251">Example</span></span>

<span data-ttu-id="4472b-252">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4472b-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="4472b-253">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4472b-253">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="4472b-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4472b-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="4472b-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="4472b-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-258">Le `recipientType` propriété de la `EmailAddressDetails` objet dans le `from` propriété est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4472b-258">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-259">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-259">Type:</span></span>

*   [<span data-ttu-id="4472b-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4472b-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4472b-261">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-261">Requirements</span></span>

|<span data-ttu-id="4472b-262">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-262">Requirement</span></span>| <span data-ttu-id="4472b-263">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-264">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-265">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-265">1.0</span></span>|
|[<span data-ttu-id="4472b-266">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-267">ReadItem</span></span>|
|[<span data-ttu-id="4472b-268">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-269">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="4472b-270">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="4472b-270">internetMessageId :String</span></span>

<span data-ttu-id="4472b-p114">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4472b-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-273">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-273">Type:</span></span>

*   <span data-ttu-id="4472b-274">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4472b-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-275">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-275">Requirements</span></span>

|<span data-ttu-id="4472b-276">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-276">Requirement</span></span>| <span data-ttu-id="4472b-277">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-278">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-279">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-279">1.0</span></span>|
|[<span data-ttu-id="4472b-280">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-281">ReadItem</span></span>|
|[<span data-ttu-id="4472b-282">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-283">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-284">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-284">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="4472b-285">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="4472b-285">itemClass :String</span></span>

<span data-ttu-id="4472b-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4472b-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="4472b-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4472b-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="4472b-290">Type</span><span class="sxs-lookup"><span data-stu-id="4472b-290">Type</span></span> | <span data-ttu-id="4472b-291">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-291">Description</span></span> | <span data-ttu-id="4472b-292">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="4472b-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="4472b-293">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="4472b-293">Appointment items</span></span> | <span data-ttu-id="4472b-294">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="4472b-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="4472b-295">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="4472b-295">Message items</span></span> | <span data-ttu-id="4472b-296">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="4472b-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="4472b-297">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="4472b-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-298">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-298">Type:</span></span>

*   <span data-ttu-id="4472b-299">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4472b-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-300">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-300">Requirements</span></span>

|<span data-ttu-id="4472b-301">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-301">Requirement</span></span>| <span data-ttu-id="4472b-302">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-303">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-304">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-304">1.0</span></span>|
|[<span data-ttu-id="4472b-305">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-306">ReadItem</span></span>|
|[<span data-ttu-id="4472b-307">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-308">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-309">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-309">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="4472b-310">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="4472b-310">(nullable) itemId :String</span></span>

<span data-ttu-id="4472b-p117">Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4472b-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-313">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="4472b-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4472b-314">Le `itemId` propriété n’est pas identique à l’identificateur d’entrée Outlook ou l’ID utilisé par l’API REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="4472b-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="4472b-315">Avant d’effectuer des appels d’API REST à l’aide de cette valeur, elle doit être convertie à l’aide de [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4472b-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4472b-316">Pour plus d’informations, voir [utiliser les API REST d’Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="4472b-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="4472b-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4472b-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-319">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-319">Type:</span></span>

*   <span data-ttu-id="4472b-320">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4472b-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-321">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-321">Requirements</span></span>

|<span data-ttu-id="4472b-322">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-322">Requirement</span></span>| <span data-ttu-id="4472b-323">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-324">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-324">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-325">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-325">1.0</span></span>|
|[<span data-ttu-id="4472b-326">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-327">ReadItem</span></span>|
|[<span data-ttu-id="4472b-328">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-329">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-330">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-330">Example</span></span>

<span data-ttu-id="4472b-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4472b-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="4472b-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="4472b-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="4472b-334">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="4472b-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="4472b-335">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4472b-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-336">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-336">Type:</span></span>

*   [<span data-ttu-id="4472b-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="4472b-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="4472b-338">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-338">Requirements</span></span>

|<span data-ttu-id="4472b-339">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-339">Requirement</span></span>| <span data-ttu-id="4472b-340">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-341">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-341">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-342">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-342">1.0</span></span>|
|[<span data-ttu-id="4472b-343">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-344">ReadItem</span></span>|
|[<span data-ttu-id="4472b-345">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-346">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-347">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-347">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="4472b-348">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="4472b-348">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="4472b-349">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4472b-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4472b-350">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-350">Read mode</span></span>

<span data-ttu-id="4472b-351">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4472b-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4472b-352">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4472b-352">Compose mode</span></span>

<span data-ttu-id="4472b-353">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4472b-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-354">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-354">Type:</span></span>

*   <span data-ttu-id="4472b-355">String | [Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="4472b-355">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-356">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-356">Requirements</span></span>

|<span data-ttu-id="4472b-357">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-357">Requirement</span></span>| <span data-ttu-id="4472b-358">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-359">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-359">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-360">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-360">1.0</span></span>|
|[<span data-ttu-id="4472b-361">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-362">ReadItem</span></span>|
|[<span data-ttu-id="4472b-363">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-364">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-365">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-365">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="4472b-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="4472b-366">normalizedSubject :String</span></span>

<span data-ttu-id="4472b-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4472b-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="4472b-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject).</span><span class="sxs-lookup"><span data-stu-id="4472b-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-371">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-371">Type:</span></span>

*   <span data-ttu-id="4472b-372">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4472b-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-373">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-373">Requirements</span></span>

|<span data-ttu-id="4472b-374">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-374">Requirement</span></span>| <span data-ttu-id="4472b-375">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-376">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-376">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-377">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-377">1.0</span></span>|
|[<span data-ttu-id="4472b-378">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-379">ReadItem</span></span>|
|[<span data-ttu-id="4472b-380">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-381">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-382">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-382">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="4472b-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="4472b-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="4472b-384">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="4472b-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-385">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-385">Type:</span></span>

*   [<span data-ttu-id="4472b-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="4472b-386">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="4472b-387">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-387">Requirements</span></span>

|<span data-ttu-id="4472b-388">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-388">Requirement</span></span>| <span data-ttu-id="4472b-389">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-390">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-390">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-391">1.3</span><span class="sxs-lookup"><span data-stu-id="4472b-391">1.3</span></span>|
|[<span data-ttu-id="4472b-392">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-393">ReadItem</span></span>|
|[<span data-ttu-id="4472b-394">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-395">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="4472b-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4472b-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="4472b-397">Fournit l’accès à un événement les participants facultatifs.</span><span class="sxs-lookup"><span data-stu-id="4472b-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="4472b-398">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="4472b-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4472b-399">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-399">Read mode</span></span>

<span data-ttu-id="4472b-400">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="4472b-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4472b-401">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4472b-401">Compose mode</span></span>

<span data-ttu-id="4472b-402">Le `optionalAttendees` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs à une réunion.</span><span class="sxs-lookup"><span data-stu-id="4472b-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-403">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-403">Type:</span></span>

*   <span data-ttu-id="4472b-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4472b-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-405">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-405">Requirements</span></span>

|<span data-ttu-id="4472b-406">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-406">Requirement</span></span>| <span data-ttu-id="4472b-407">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-408">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-408">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-409">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-409">1.0</span></span>|
|[<span data-ttu-id="4472b-410">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-411">ReadItem</span></span>|
|[<span data-ttu-id="4472b-412">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-413">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-414">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-414">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="4472b-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4472b-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="4472b-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4472b-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-418">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-418">Type:</span></span>

*   [<span data-ttu-id="4472b-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4472b-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4472b-420">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-420">Requirements</span></span>

|<span data-ttu-id="4472b-421">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-421">Requirement</span></span>| <span data-ttu-id="4472b-422">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-423">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-423">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-424">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-424">1.0</span></span>|
|[<span data-ttu-id="4472b-425">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-426">ReadItem</span></span>|
|[<span data-ttu-id="4472b-427">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-428">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-429">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-429">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="4472b-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4472b-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="4472b-431">Fournit l’accès à des participants obligatoires d’un événement.</span><span class="sxs-lookup"><span data-stu-id="4472b-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="4472b-432">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="4472b-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4472b-433">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-433">Read mode</span></span>

<span data-ttu-id="4472b-434">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="4472b-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4472b-435">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4472b-435">Compose mode</span></span>

<span data-ttu-id="4472b-436">Le `requiredAttendees` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les participants obligatoires à une réunion.</span><span class="sxs-lookup"><span data-stu-id="4472b-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-437">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-437">Type:</span></span>

*   <span data-ttu-id="4472b-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4472b-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-439">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-439">Requirements</span></span>

|<span data-ttu-id="4472b-440">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-440">Requirement</span></span>| <span data-ttu-id="4472b-441">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-442">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-443">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-443">1.0</span></span>|
|[<span data-ttu-id="4472b-444">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-445">ReadItem</span></span>|
|[<span data-ttu-id="4472b-446">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-447">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-448">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-448">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="4472b-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4472b-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="4472b-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4472b-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="4472b-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="4472b-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-454">Le `recipientType` propriété de la `EmailAddressDetails` objet dans le `sender` propriété est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4472b-454">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-455">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-455">Type:</span></span>

*   [<span data-ttu-id="4472b-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4472b-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4472b-457">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-457">Requirements</span></span>

|<span data-ttu-id="4472b-458">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-458">Requirement</span></span>| <span data-ttu-id="4472b-459">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-460">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-460">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-461">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-461">1.0</span></span>|
|[<span data-ttu-id="4472b-462">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-463">ReadItem</span></span>|
|[<span data-ttu-id="4472b-464">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-465">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-466">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-466">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="4472b-467">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="4472b-467">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="4472b-468">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4472b-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="4472b-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="4472b-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4472b-471">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-471">Read mode</span></span>

<span data-ttu-id="4472b-472">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="4472b-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4472b-473">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4472b-473">Compose mode</span></span>

<span data-ttu-id="4472b-474">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4472b-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="4472b-475">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="4472b-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-476">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-476">Type:</span></span>

*   <span data-ttu-id="4472b-477">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="4472b-477">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-478">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-478">Requirements</span></span>

|<span data-ttu-id="4472b-479">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-479">Requirement</span></span>| <span data-ttu-id="4472b-480">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-481">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-481">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-482">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-482">1.0</span></span>|
|[<span data-ttu-id="4472b-483">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-484">ReadItem</span></span>|
|[<span data-ttu-id="4472b-485">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-486">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-487">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-487">Example</span></span>

<span data-ttu-id="4472b-488">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="4472b-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="4472b-489">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4472b-489">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="4472b-490">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4472b-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="4472b-491">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="4472b-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4472b-492">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-492">Read mode</span></span>

<span data-ttu-id="4472b-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="4472b-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="4472b-495">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4472b-495">Compose mode</span></span>

<span data-ttu-id="4472b-496">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="4472b-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4472b-497">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-497">Type:</span></span>

*   <span data-ttu-id="4472b-498">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4472b-498">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-499">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-499">Requirements</span></span>

|<span data-ttu-id="4472b-500">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-500">Requirement</span></span>| <span data-ttu-id="4472b-501">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-502">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-502">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-503">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-503">1.0</span></span>|
|[<span data-ttu-id="4472b-504">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-505">ReadItem</span></span>|
|[<span data-ttu-id="4472b-506">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-507">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="4472b-508">à : Array. <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4472b-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="4472b-509">Permet d’accéder aux destinataires de la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="4472b-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="4472b-510">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="4472b-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4472b-511">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-511">Read mode</span></span>

<span data-ttu-id="4472b-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="4472b-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="4472b-514">Mode composition</span><span class="sxs-lookup"><span data-stu-id="4472b-514">Compose mode</span></span>

<span data-ttu-id="4472b-515">Le `to` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="4472b-515">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="4472b-516">Type :</span><span class="sxs-lookup"><span data-stu-id="4472b-516">Type:</span></span>

*   <span data-ttu-id="4472b-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4472b-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-518">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-518">Requirements</span></span>

|<span data-ttu-id="4472b-519">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-519">Requirement</span></span>| <span data-ttu-id="4472b-520">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-521">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-522">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-522">1.0</span></span>|
|[<span data-ttu-id="4472b-523">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-524">ReadItem</span></span>|
|[<span data-ttu-id="4472b-525">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-526">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-527">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-527">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="4472b-528">Méthodes</span><span class="sxs-lookup"><span data-stu-id="4472b-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="4472b-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4472b-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4472b-530">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="4472b-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4472b-531">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="4472b-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="4472b-532">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="4472b-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4472b-533">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4472b-533">Parameters:</span></span>

|<span data-ttu-id="4472b-534">Nom</span><span class="sxs-lookup"><span data-stu-id="4472b-534">Name</span></span>| <span data-ttu-id="4472b-535">Type</span><span class="sxs-lookup"><span data-stu-id="4472b-535">Type</span></span>| <span data-ttu-id="4472b-536">Attributs</span><span class="sxs-lookup"><span data-stu-id="4472b-536">Attributes</span></span>| <span data-ttu-id="4472b-537">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="4472b-538">String</span><span class="sxs-lookup"><span data-stu-id="4472b-538">String</span></span>||<span data-ttu-id="4472b-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="4472b-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4472b-541">String</span><span class="sxs-lookup"><span data-stu-id="4472b-541">String</span></span>||<span data-ttu-id="4472b-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4472b-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4472b-544">Objet</span><span class="sxs-lookup"><span data-stu-id="4472b-544">Object</span></span>| <span data-ttu-id="4472b-545">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-545">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-546">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4472b-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4472b-547">Objet</span><span class="sxs-lookup"><span data-stu-id="4472b-547">Object</span></span>| <span data-ttu-id="4472b-548">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-548">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-549">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4472b-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4472b-550">fonction</span><span class="sxs-lookup"><span data-stu-id="4472b-550">function</span></span>| <span data-ttu-id="4472b-551">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-551">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-552">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4472b-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4472b-553">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4472b-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4472b-554">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="4472b-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4472b-555">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4472b-555">Errors</span></span>

| <span data-ttu-id="4472b-556">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4472b-556">Error code</span></span> | <span data-ttu-id="4472b-557">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="4472b-558">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="4472b-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="4472b-559">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="4472b-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4472b-560">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4472b-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4472b-561">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-561">Requirements</span></span>

|<span data-ttu-id="4472b-562">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-562">Requirement</span></span>| <span data-ttu-id="4472b-563">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-564">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-565">1.1</span><span class="sxs-lookup"><span data-stu-id="4472b-565">1.1</span></span>|
|[<span data-ttu-id="4472b-566">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4472b-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="4472b-568">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-569">Composition</span><span class="sxs-lookup"><span data-stu-id="4472b-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-570">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-570">Example</span></span>

```
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="4472b-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4472b-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4472b-572">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4472b-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="4472b-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4472b-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="4472b-576">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="4472b-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="4472b-577">Si votre complément Office est en cours d’exécution dans Outlook Web App, le `addItemAttachmentAsync` méthode permettre joindre des éléments à des éléments autres que l’élément à modifier ; Toutefois, cela n’est pas pris en charge et n’est pas recommandée.</span><span class="sxs-lookup"><span data-stu-id="4472b-577">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4472b-578">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4472b-578">Parameters:</span></span>

|<span data-ttu-id="4472b-579">Nom</span><span class="sxs-lookup"><span data-stu-id="4472b-579">Name</span></span>| <span data-ttu-id="4472b-580">Type</span><span class="sxs-lookup"><span data-stu-id="4472b-580">Type</span></span>| <span data-ttu-id="4472b-581">Attributs</span><span class="sxs-lookup"><span data-stu-id="4472b-581">Attributes</span></span>| <span data-ttu-id="4472b-582">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="4472b-583">String</span><span class="sxs-lookup"><span data-stu-id="4472b-583">String</span></span>||<span data-ttu-id="4472b-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4472b-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4472b-586">String</span><span class="sxs-lookup"><span data-stu-id="4472b-586">String</span></span>||<span data-ttu-id="4472b-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4472b-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4472b-589">Objet</span><span class="sxs-lookup"><span data-stu-id="4472b-589">Object</span></span>| <span data-ttu-id="4472b-590">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-590">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-591">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4472b-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4472b-592">Objet</span><span class="sxs-lookup"><span data-stu-id="4472b-592">Object</span></span>| <span data-ttu-id="4472b-593">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-593">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-594">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4472b-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4472b-595">fonction</span><span class="sxs-lookup"><span data-stu-id="4472b-595">function</span></span>| <span data-ttu-id="4472b-596">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-596">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-597">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4472b-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4472b-598">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4472b-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4472b-599">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="4472b-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4472b-600">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4472b-600">Errors</span></span>

| <span data-ttu-id="4472b-601">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4472b-601">Error code</span></span> | <span data-ttu-id="4472b-602">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4472b-603">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4472b-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4472b-604">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-604">Requirements</span></span>

|<span data-ttu-id="4472b-605">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-605">Requirement</span></span>| <span data-ttu-id="4472b-606">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-607">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-607">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-608">1.1</span><span class="sxs-lookup"><span data-stu-id="4472b-608">1.1</span></span>|
|[<span data-ttu-id="4472b-609">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4472b-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="4472b-611">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-612">Composition</span><span class="sxs-lookup"><span data-stu-id="4472b-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-613">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-613">Example</span></span>

<span data-ttu-id="4472b-614">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="4472b-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="4472b-615">close()</span><span class="sxs-lookup"><span data-stu-id="4472b-615">close()</span></span>

<span data-ttu-id="4472b-616">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="4472b-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="4472b-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="4472b-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-619">Dans Outlook sur le web, si l’élément est un rendez-vous et il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, supprimer ou annuler même si aucune modification de l’élément depuis le dernier enregistrée.</span><span class="sxs-lookup"><span data-stu-id="4472b-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="4472b-620">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="4472b-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-621">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-621">Requirements</span></span>

|<span data-ttu-id="4472b-622">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-622">Requirement</span></span>| <span data-ttu-id="4472b-623">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-624">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-624">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-625">1.3</span><span class="sxs-lookup"><span data-stu-id="4472b-625">1.3</span></span>|
|[<span data-ttu-id="4472b-626">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-627">Restricted</span><span class="sxs-lookup"><span data-stu-id="4472b-627">Restricted</span></span>|
|[<span data-ttu-id="4472b-628">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-629">Composition</span><span class="sxs-lookup"><span data-stu-id="4472b-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="4472b-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="4472b-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="4472b-631">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4472b-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-632">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4472b-632">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4472b-633">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="4472b-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4472b-634">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="4472b-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="4472b-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="4472b-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4472b-638">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4472b-638">Parameters:</span></span>

|<span data-ttu-id="4472b-639">Nom</span><span class="sxs-lookup"><span data-stu-id="4472b-639">Name</span></span>| <span data-ttu-id="4472b-640">Type</span><span class="sxs-lookup"><span data-stu-id="4472b-640">Type</span></span>| <span data-ttu-id="4472b-641">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="4472b-642">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4472b-642">String &#124; Object</span></span>| |<span data-ttu-id="4472b-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4472b-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4472b-645">**OU**</span><span class="sxs-lookup"><span data-stu-id="4472b-645">**OR**</span></span><br/><span data-ttu-id="4472b-p140">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="4472b-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4472b-648">String</span><span class="sxs-lookup"><span data-stu-id="4472b-648">String</span></span> | <span data-ttu-id="4472b-649">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-649">&lt;optional&gt;</span></span> | <span data-ttu-id="4472b-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4472b-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4472b-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4472b-653">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-653">&lt;optional&gt;</span></span> | <span data-ttu-id="4472b-654">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="4472b-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4472b-655">String</span><span class="sxs-lookup"><span data-stu-id="4472b-655">String</span></span> | | <span data-ttu-id="4472b-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="4472b-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4472b-658">String</span><span class="sxs-lookup"><span data-stu-id="4472b-658">String</span></span> | | <span data-ttu-id="4472b-659">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4472b-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4472b-660">String</span><span class="sxs-lookup"><span data-stu-id="4472b-660">String</span></span> | | <span data-ttu-id="4472b-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="4472b-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4472b-663">String</span><span class="sxs-lookup"><span data-stu-id="4472b-663">String</span></span> | | <span data-ttu-id="4472b-p144">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4472b-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4472b-667">function</span><span class="sxs-lookup"><span data-stu-id="4472b-667">function</span></span> | <span data-ttu-id="4472b-668">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-668">&lt;optional&gt;</span></span> | <span data-ttu-id="4472b-669">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4472b-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4472b-670">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-670">Requirements</span></span>

|<span data-ttu-id="4472b-671">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-671">Requirement</span></span>| <span data-ttu-id="4472b-672">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-673">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-673">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-674">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-674">1.0</span></span>|
|[<span data-ttu-id="4472b-675">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-676">ReadItem</span></span>|
|[<span data-ttu-id="4472b-677">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-678">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4472b-679">Exemples</span><span class="sxs-lookup"><span data-stu-id="4472b-679">Examples</span></span>

<span data-ttu-id="4472b-680">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="4472b-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="4472b-681">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="4472b-681">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="4472b-682">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="4472b-682">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4472b-683">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="4472b-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4472b-684">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4472b-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4472b-685">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="4472b-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="4472b-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="4472b-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="4472b-687">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4472b-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-688">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4472b-688">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4472b-689">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="4472b-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4472b-690">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="4472b-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="4472b-p145">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="4472b-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4472b-694">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4472b-694">Parameters:</span></span>

|<span data-ttu-id="4472b-695">Nom</span><span class="sxs-lookup"><span data-stu-id="4472b-695">Name</span></span>| <span data-ttu-id="4472b-696">Type</span><span class="sxs-lookup"><span data-stu-id="4472b-696">Type</span></span>| <span data-ttu-id="4472b-697">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="4472b-698">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4472b-698">String &#124; Object</span></span>| | <span data-ttu-id="4472b-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4472b-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4472b-701">**OU**</span><span class="sxs-lookup"><span data-stu-id="4472b-701">**OR**</span></span><br/><span data-ttu-id="4472b-p147">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="4472b-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4472b-704">String</span><span class="sxs-lookup"><span data-stu-id="4472b-704">String</span></span> | <span data-ttu-id="4472b-705">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-705">&lt;optional&gt;</span></span> | <span data-ttu-id="4472b-p148">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4472b-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4472b-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4472b-709">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-709">&lt;optional&gt;</span></span> | <span data-ttu-id="4472b-710">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="4472b-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4472b-711">String</span><span class="sxs-lookup"><span data-stu-id="4472b-711">String</span></span> | | <span data-ttu-id="4472b-p149">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="4472b-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4472b-714">String</span><span class="sxs-lookup"><span data-stu-id="4472b-714">String</span></span> | | <span data-ttu-id="4472b-715">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4472b-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4472b-716">String</span><span class="sxs-lookup"><span data-stu-id="4472b-716">String</span></span> | | <span data-ttu-id="4472b-p150">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="4472b-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4472b-719">String</span><span class="sxs-lookup"><span data-stu-id="4472b-719">String</span></span> | | <span data-ttu-id="4472b-p151">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4472b-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4472b-723">function</span><span class="sxs-lookup"><span data-stu-id="4472b-723">function</span></span> | <span data-ttu-id="4472b-724">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-724">&lt;optional&gt;</span></span> | <span data-ttu-id="4472b-725">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4472b-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4472b-726">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-726">Requirements</span></span>

|<span data-ttu-id="4472b-727">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-727">Requirement</span></span>| <span data-ttu-id="4472b-728">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-729">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-729">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-730">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-730">1.0</span></span>|
|[<span data-ttu-id="4472b-731">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-732">ReadItem</span></span>|
|[<span data-ttu-id="4472b-733">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-734">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4472b-735">Exemples</span><span class="sxs-lookup"><span data-stu-id="4472b-735">Examples</span></span>

<span data-ttu-id="4472b-736">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="4472b-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="4472b-737">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="4472b-737">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="4472b-738">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="4472b-738">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4472b-739">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="4472b-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4472b-740">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4472b-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4472b-741">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="4472b-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="4472b-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="4472b-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="4472b-743">Obtient les entités trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4472b-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-744">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4472b-744">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-745">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-745">Requirements</span></span>

|<span data-ttu-id="4472b-746">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-746">Requirement</span></span>| <span data-ttu-id="4472b-747">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-748">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-748">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-749">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-749">1.0</span></span>|
|[<span data-ttu-id="4472b-750">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-751">ReadItem</span></span>|
|[<span data-ttu-id="4472b-752">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-753">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4472b-754">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4472b-754">Returns:</span></span>

<span data-ttu-id="4472b-755">Type : [Entities](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="4472b-755">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="4472b-756">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-756">Example</span></span>

<span data-ttu-id="4472b-757">L’exemple suivant accède aux entités de contacts dans le corps de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="4472b-757">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="4472b-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4472b-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4472b-759">Obtient un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4472b-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-760">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4472b-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4472b-761">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4472b-761">Parameters:</span></span>

|<span data-ttu-id="4472b-762">Nom</span><span class="sxs-lookup"><span data-stu-id="4472b-762">Name</span></span>| <span data-ttu-id="4472b-763">Type</span><span class="sxs-lookup"><span data-stu-id="4472b-763">Type</span></span>| <span data-ttu-id="4472b-764">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="4472b-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="4472b-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="4472b-766">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="4472b-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4472b-767">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-767">Requirements</span></span>

|<span data-ttu-id="4472b-768">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-768">Requirement</span></span>| <span data-ttu-id="4472b-769">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-770">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-770">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-771">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-771">1.0</span></span>|
|[<span data-ttu-id="4472b-772">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-773">Restricted</span><span class="sxs-lookup"><span data-stu-id="4472b-773">Restricted</span></span>|
|[<span data-ttu-id="4472b-774">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-775">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4472b-776">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4472b-776">Returns:</span></span>

<span data-ttu-id="4472b-777">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="4472b-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="4472b-778">Si aucune entité du type spécifié n’est présentes dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="4472b-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="4472b-779">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="4472b-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="4472b-780">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="4472b-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="4472b-781">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="4472b-781">Value of `entityType`</span></span> | <span data-ttu-id="4472b-782">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="4472b-782">Type of objects in returned array</span></span> | <span data-ttu-id="4472b-783">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="4472b-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="4472b-784">String</span><span class="sxs-lookup"><span data-stu-id="4472b-784">String</span></span> | <span data-ttu-id="4472b-785">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4472b-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="4472b-786">Contact</span><span class="sxs-lookup"><span data-stu-id="4472b-786">Contact</span></span> | <span data-ttu-id="4472b-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4472b-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="4472b-788">String</span><span class="sxs-lookup"><span data-stu-id="4472b-788">String</span></span> | <span data-ttu-id="4472b-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4472b-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="4472b-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4472b-790">MeetingSuggestion</span></span> | <span data-ttu-id="4472b-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4472b-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="4472b-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4472b-792">PhoneNumber</span></span> | <span data-ttu-id="4472b-793">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4472b-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="4472b-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4472b-794">TaskSuggestion</span></span> | <span data-ttu-id="4472b-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4472b-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="4472b-796">String</span><span class="sxs-lookup"><span data-stu-id="4472b-796">String</span></span> | <span data-ttu-id="4472b-797">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4472b-797">**Restricted**</span></span> |

<span data-ttu-id="4472b-798">Type : Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4472b-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="4472b-799">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-799">Example</span></span>

<span data-ttu-id="4472b-800">L’exemple suivant montre comment accéder à un tableau de chaînes qui représentent les adresses postales dans le corps de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="4472b-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="4472b-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4472b-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4472b-802">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4472b-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-803">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4472b-803">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4472b-804">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="4472b-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4472b-805">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4472b-805">Parameters:</span></span>

|<span data-ttu-id="4472b-806">Nom</span><span class="sxs-lookup"><span data-stu-id="4472b-806">Name</span></span>| <span data-ttu-id="4472b-807">Type</span><span class="sxs-lookup"><span data-stu-id="4472b-807">Type</span></span>| <span data-ttu-id="4472b-808">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4472b-809">String</span><span class="sxs-lookup"><span data-stu-id="4472b-809">String</span></span>|<span data-ttu-id="4472b-810">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="4472b-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4472b-811">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-811">Requirements</span></span>

|<span data-ttu-id="4472b-812">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-812">Requirement</span></span>| <span data-ttu-id="4472b-813">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-814">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-814">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-815">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-815">1.0</span></span>|
|[<span data-ttu-id="4472b-816">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-817">ReadItem</span></span>|
|[<span data-ttu-id="4472b-818">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-819">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4472b-820">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4472b-820">Returns:</span></span>

<span data-ttu-id="4472b-p153">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="4472b-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="4472b-823">Type : Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4472b-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="4472b-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4472b-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="4472b-825">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4472b-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-826">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4472b-826">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4472b-p154">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="4472b-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4472b-830">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="4472b-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4472b-831">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4472b-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4472b-p155">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4472b-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4472b-835">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-835">Requirements</span></span>

|<span data-ttu-id="4472b-836">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-836">Requirement</span></span>| <span data-ttu-id="4472b-837">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-838">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-838">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-839">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-839">1.0</span></span>|
|[<span data-ttu-id="4472b-840">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-841">ReadItem</span></span>|
|[<span data-ttu-id="4472b-842">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-843">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4472b-844">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4472b-844">Returns:</span></span>

<span data-ttu-id="4472b-p156">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="4472b-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="4472b-847">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4472b-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4472b-848">Object</span><span class="sxs-lookup"><span data-stu-id="4472b-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4472b-849">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-849">Example</span></span>

<span data-ttu-id="4472b-850">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="4472b-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="4472b-851">getregexmatchesbyname (Name) → (nullable) {tableau. < chaîne >}</span><span class="sxs-lookup"><span data-stu-id="4472b-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="4472b-852">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4472b-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-853">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4472b-853">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4472b-854">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="4472b-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="4472b-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="4472b-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4472b-857">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4472b-857">Parameters:</span></span>

|<span data-ttu-id="4472b-858">Nom</span><span class="sxs-lookup"><span data-stu-id="4472b-858">Name</span></span>| <span data-ttu-id="4472b-859">Type</span><span class="sxs-lookup"><span data-stu-id="4472b-859">Type</span></span>| <span data-ttu-id="4472b-860">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4472b-861">String</span><span class="sxs-lookup"><span data-stu-id="4472b-861">String</span></span>|<span data-ttu-id="4472b-862">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="4472b-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4472b-863">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-863">Requirements</span></span>

|<span data-ttu-id="4472b-864">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-864">Requirement</span></span>| <span data-ttu-id="4472b-865">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-866">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-866">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-867">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-867">1.0</span></span>|
|[<span data-ttu-id="4472b-868">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-869">ReadItem</span></span>|
|[<span data-ttu-id="4472b-870">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-871">Lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4472b-872">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4472b-872">Returns:</span></span>

<span data-ttu-id="4472b-873">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="4472b-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="4472b-874">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4472b-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4472b-875">< Chaîne > tableau.</span><span class="sxs-lookup"><span data-stu-id="4472b-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4472b-876">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-876">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="4472b-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="4472b-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="4472b-878">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="4472b-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="4472b-p158">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="4472b-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4472b-881">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4472b-881">Parameters:</span></span>

|<span data-ttu-id="4472b-882">Nom</span><span class="sxs-lookup"><span data-stu-id="4472b-882">Name</span></span>| <span data-ttu-id="4472b-883">Type</span><span class="sxs-lookup"><span data-stu-id="4472b-883">Type</span></span>| <span data-ttu-id="4472b-884">Attributs</span><span class="sxs-lookup"><span data-stu-id="4472b-884">Attributes</span></span>| <span data-ttu-id="4472b-885">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="4472b-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4472b-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="4472b-p159">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="4472b-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="4472b-890">Object</span><span class="sxs-lookup"><span data-stu-id="4472b-890">Object</span></span>| <span data-ttu-id="4472b-891">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-891">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-892">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4472b-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4472b-893">Objet</span><span class="sxs-lookup"><span data-stu-id="4472b-893">Object</span></span>| <span data-ttu-id="4472b-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-894">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-895">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4472b-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4472b-896">fonction</span><span class="sxs-lookup"><span data-stu-id="4472b-896">function</span></span>||<span data-ttu-id="4472b-897">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4472b-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4472b-898">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="4472b-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="4472b-899">Pour accéder à la propriété source provenant de la sélection, appelez `asyncResult.value.sourceProperty`, qui sera soit `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="4472b-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4472b-900">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-900">Requirements</span></span>

|<span data-ttu-id="4472b-901">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-901">Requirement</span></span>| <span data-ttu-id="4472b-902">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-903">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-903">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-904">1.2</span><span class="sxs-lookup"><span data-stu-id="4472b-904">1.2</span></span>|
|[<span data-ttu-id="4472b-905">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4472b-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="4472b-907">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-908">Composition</span><span class="sxs-lookup"><span data-stu-id="4472b-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4472b-909">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4472b-909">Returns:</span></span>

<span data-ttu-id="4472b-910">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="4472b-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="4472b-911">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4472b-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4472b-912">String</span><span class="sxs-lookup"><span data-stu-id="4472b-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4472b-913">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="4472b-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4472b-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="4472b-915">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="4472b-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="4472b-p161">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="4472b-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4472b-919">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4472b-919">Parameters:</span></span>

|<span data-ttu-id="4472b-920">Nom</span><span class="sxs-lookup"><span data-stu-id="4472b-920">Name</span></span>| <span data-ttu-id="4472b-921">Type</span><span class="sxs-lookup"><span data-stu-id="4472b-921">Type</span></span>| <span data-ttu-id="4472b-922">Attributs</span><span class="sxs-lookup"><span data-stu-id="4472b-922">Attributes</span></span>| <span data-ttu-id="4472b-923">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4472b-924">function</span><span class="sxs-lookup"><span data-stu-id="4472b-924">function</span></span>||<span data-ttu-id="4472b-925">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4472b-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4472b-926">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4472b-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4472b-927">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées de l’élément et enregistrer les modifications à la propriété personnalisée est redéfinie sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="4472b-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="4472b-928">Objet</span><span class="sxs-lookup"><span data-stu-id="4472b-928">Object</span></span>| <span data-ttu-id="4472b-929">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-929">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-930">Les développeurs peuvent fournir n’importe quel objet qu’ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4472b-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="4472b-931">Cet objet est accessible par le `asyncResult.asyncContext` propriété dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="4472b-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4472b-932">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-932">Requirements</span></span>

|<span data-ttu-id="4472b-933">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-933">Requirement</span></span>| <span data-ttu-id="4472b-934">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-935">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-936">1.0</span><span class="sxs-lookup"><span data-stu-id="4472b-936">1.0</span></span>|
|[<span data-ttu-id="4472b-937">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4472b-938">ReadItem</span></span>|
|[<span data-ttu-id="4472b-939">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-940">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4472b-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-941">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-941">Example</span></span>

<span data-ttu-id="4472b-p164">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="4472b-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="4472b-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4472b-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="4472b-946">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4472b-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="4472b-p165">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="4472b-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4472b-951">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4472b-951">Parameters:</span></span>

|<span data-ttu-id="4472b-952">Nom</span><span class="sxs-lookup"><span data-stu-id="4472b-952">Name</span></span>| <span data-ttu-id="4472b-953">Type</span><span class="sxs-lookup"><span data-stu-id="4472b-953">Type</span></span>| <span data-ttu-id="4472b-954">Attributs</span><span class="sxs-lookup"><span data-stu-id="4472b-954">Attributes</span></span>| <span data-ttu-id="4472b-955">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="4472b-956">String</span><span class="sxs-lookup"><span data-stu-id="4472b-956">String</span></span>||<span data-ttu-id="4472b-p166">Identificateur de la pièce jointe à supprimer. La longueur maximale de la chaîne est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4472b-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="4472b-959">Objet</span><span class="sxs-lookup"><span data-stu-id="4472b-959">Object</span></span>| <span data-ttu-id="4472b-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-960">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-961">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4472b-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4472b-962">Objet</span><span class="sxs-lookup"><span data-stu-id="4472b-962">Object</span></span>| <span data-ttu-id="4472b-963">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-963">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-964">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4472b-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4472b-965">fonction</span><span class="sxs-lookup"><span data-stu-id="4472b-965">function</span></span>| <span data-ttu-id="4472b-966">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-966">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-967">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4472b-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4472b-968">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="4472b-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4472b-969">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4472b-969">Errors</span></span>

| <span data-ttu-id="4472b-970">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4472b-970">Error code</span></span> | <span data-ttu-id="4472b-971">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="4472b-972">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="4472b-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4472b-973">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-973">Requirements</span></span>

|<span data-ttu-id="4472b-974">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-974">Requirement</span></span>| <span data-ttu-id="4472b-975">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-976">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-976">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-977">1.1</span><span class="sxs-lookup"><span data-stu-id="4472b-977">1.1</span></span>|
|[<span data-ttu-id="4472b-978">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4472b-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="4472b-980">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-981">Composition</span><span class="sxs-lookup"><span data-stu-id="4472b-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-982">Exemple</span><span class="sxs-lookup"><span data-stu-id="4472b-982">Example</span></span>

<span data-ttu-id="4472b-983">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="4472b-983">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="4472b-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4472b-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="4472b-985">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4472b-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="4472b-p167">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="4472b-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-989">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` pour utiliser avec EWS ou l’API REST, gardez à l’esprit que quand Outlook est en mode mis en cache, il peut prendre un certain temps avant que l’élément est réellement synchronisé avec le serveur.</span><span class="sxs-lookup"><span data-stu-id="4472b-989">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="4472b-990">Jusqu'à ce que l’élément est synchronisé, à l’aide de la `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="4472b-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="4472b-p169">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="4472b-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="4472b-994">Les clients suivants ont un comportement différent pour `saveAsync` de rendez-vous dans le mode de composition :</span><span class="sxs-lookup"><span data-stu-id="4472b-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="4472b-995">Mac Outlook ne gère pas `saveAsync` d’une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="4472b-995">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="4472b-996">Appel de `saveAsync` d’une réunion dans Outlook Mac renverra une erreur.</span><span class="sxs-lookup"><span data-stu-id="4472b-996">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="4472b-997">Outlook sur le web toujours envoie une invitation ou mettre à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="4472b-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4472b-998">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4472b-998">Parameters:</span></span>

|<span data-ttu-id="4472b-999">Nom</span><span class="sxs-lookup"><span data-stu-id="4472b-999">Name</span></span>| <span data-ttu-id="4472b-1000">Type</span><span class="sxs-lookup"><span data-stu-id="4472b-1000">Type</span></span>| <span data-ttu-id="4472b-1001">Attributs</span><span class="sxs-lookup"><span data-stu-id="4472b-1001">Attributes</span></span>| <span data-ttu-id="4472b-1002">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="4472b-1003">Objet</span><span class="sxs-lookup"><span data-stu-id="4472b-1003">Object</span></span>| <span data-ttu-id="4472b-1004">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-1005">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4472b-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4472b-1006">Objet</span><span class="sxs-lookup"><span data-stu-id="4472b-1006">Object</span></span>| <span data-ttu-id="4472b-1007">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-1008">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4472b-1008">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="4472b-1009">fonction</span><span class="sxs-lookup"><span data-stu-id="4472b-1009">function</span></span>||<span data-ttu-id="4472b-1010">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4472b-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4472b-1011">En cas de réussite, l’identificateur d’élément est fournie dans le `asyncResult.value` propriété.</span><span class="sxs-lookup"><span data-stu-id="4472b-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4472b-1012">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-1012">Requirements</span></span>

|<span data-ttu-id="4472b-1013">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-1013">Requirement</span></span>| <span data-ttu-id="4472b-1014">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-1015">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-1015">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="4472b-1016">1.3</span></span>|
|[<span data-ttu-id="4472b-1017">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4472b-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="4472b-1019">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-1020">Composition</span><span class="sxs-lookup"><span data-stu-id="4472b-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4472b-1021">範例</span><span class="sxs-lookup"><span data-stu-id="4472b-1021">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="4472b-p171">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="4472b-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="4472b-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="4472b-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="4472b-1025">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4472b-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="4472b-p172">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="4472b-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4472b-1029">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="4472b-1029">Parameters:</span></span>

|<span data-ttu-id="4472b-1030">Nom</span><span class="sxs-lookup"><span data-stu-id="4472b-1030">Name</span></span>| <span data-ttu-id="4472b-1031">Type</span><span class="sxs-lookup"><span data-stu-id="4472b-1031">Type</span></span>| <span data-ttu-id="4472b-1032">Attributs</span><span class="sxs-lookup"><span data-stu-id="4472b-1032">Attributes</span></span>| <span data-ttu-id="4472b-1033">Description</span><span class="sxs-lookup"><span data-stu-id="4472b-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4472b-1034">String</span><span class="sxs-lookup"><span data-stu-id="4472b-1034">String</span></span>||<span data-ttu-id="4472b-p173">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="4472b-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="4472b-1038">Objet</span><span class="sxs-lookup"><span data-stu-id="4472b-1038">Object</span></span>| <span data-ttu-id="4472b-1039">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-1040">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4472b-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4472b-1041">Objet</span><span class="sxs-lookup"><span data-stu-id="4472b-1041">Object</span></span>| <span data-ttu-id="4472b-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-1043">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4472b-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="4472b-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4472b-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="4472b-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4472b-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="4472b-p174">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="4472b-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="4472b-p175">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="4472b-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="4472b-1050">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="4472b-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="4472b-1051">fonction</span><span class="sxs-lookup"><span data-stu-id="4472b-1051">function</span></span>||<span data-ttu-id="4472b-1052">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4472b-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4472b-1053">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4472b-1053">Requirements</span></span>

|<span data-ttu-id="4472b-1054">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4472b-1054">Requirement</span></span>| <span data-ttu-id="4472b-1055">Valeur</span><span class="sxs-lookup"><span data-stu-id="4472b-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="4472b-1056">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4472b-1056">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4472b-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="4472b-1057">1.2</span></span>|
|[<span data-ttu-id="4472b-1058">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4472b-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4472b-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4472b-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="4472b-1060">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4472b-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4472b-1061">Composition</span><span class="sxs-lookup"><span data-stu-id="4472b-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4472b-1062">Exemples</span><span class="sxs-lookup"><span data-stu-id="4472b-1062">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```