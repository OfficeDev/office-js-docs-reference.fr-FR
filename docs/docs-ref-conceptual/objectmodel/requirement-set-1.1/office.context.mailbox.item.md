
# <a name="item"></a><span data-ttu-id="75a2d-101">élément</span><span class="sxs-lookup"><span data-stu-id="75a2d-101">item</span></span>

### <span data-ttu-id="75a2d-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="75a2d-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="75a2d-p102">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="75a2d-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-106">Requirements</span></span>

|<span data-ttu-id="75a2d-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-107">Requirement</span></span>| <span data-ttu-id="75a2d-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-110">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-110">1.0</span></span>|
|[<span data-ttu-id="75a2d-111">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-112">Restricted</span><span class="sxs-lookup"><span data-stu-id="75a2d-112">Restricted</span></span>|
|[<span data-ttu-id="75a2d-113">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-114">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="75a2d-115">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-115">Example</span></span>

<span data-ttu-id="75a2d-116">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="75a2d-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
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

### <a name="members"></a><span data-ttu-id="75a2d-117">Membres</span><span class="sxs-lookup"><span data-stu-id="75a2d-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="75a2d-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="75a2d-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="75a2d-p103">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-121">Certains types de fichiers bloqués par Outlook en raison de problèmes de sécurité potentiels et ne sont donc pas retournés.</span><span class="sxs-lookup"><span data-stu-id="75a2d-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="75a2d-122">Pour plus d’informations, voir les [pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="75a2d-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-123">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-123">Type:</span></span>

*   <span data-ttu-id="75a2d-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="75a2d-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-125">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-125">Requirements</span></span>

|<span data-ttu-id="75a2d-126">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-126">Requirement</span></span>| <span data-ttu-id="75a2d-127">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-128">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-128">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-129">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-129">1.0</span></span>|
|[<span data-ttu-id="75a2d-130">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-131">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-132">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-133">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-134">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-134">Example</span></span>

<span data-ttu-id="75a2d-135">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="75a2d-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="75a2d-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="75a2d-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="75a2d-137">Obtient un objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires des Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="75a2d-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="75a2d-138">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="75a2d-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-139">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-139">Type:</span></span>

*   [<span data-ttu-id="75a2d-140">Destinataires</span><span class="sxs-lookup"><span data-stu-id="75a2d-140">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="75a2d-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-141">Requirements</span></span>

|<span data-ttu-id="75a2d-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-142">Requirement</span></span>| <span data-ttu-id="75a2d-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-145">1.1</span><span class="sxs-lookup"><span data-stu-id="75a2d-145">1.1</span></span>|
|[<span data-ttu-id="75a2d-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-147">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-149">Composition</span><span class="sxs-lookup"><span data-stu-id="75a2d-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="75a2d-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="75a2d-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="75a2d-152">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="75a2d-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-153">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-153">Type:</span></span>

*   [<span data-ttu-id="75a2d-154">Corps</span><span class="sxs-lookup"><span data-stu-id="75a2d-154">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="75a2d-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-155">Requirements</span></span>

|<span data-ttu-id="75a2d-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-156">Requirement</span></span>| <span data-ttu-id="75a2d-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-159">1.1</span><span class="sxs-lookup"><span data-stu-id="75a2d-159">1.1</span></span>|
|[<span data-ttu-id="75a2d-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-161">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="75a2d-164">cc : tableau. <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="75a2d-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="75a2d-165">Permet d’accéder aux destinataires Cc (copie carbone) d’un message.</span><span class="sxs-lookup"><span data-stu-id="75a2d-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="75a2d-166">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="75a2d-167">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-167">Read mode</span></span>

<span data-ttu-id="75a2d-p107">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="75a2d-170">Mode composition</span><span class="sxs-lookup"><span data-stu-id="75a2d-170">Compose mode</span></span>

<span data-ttu-id="75a2d-171">Le `cc` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="75a2d-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-172">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-172">Type:</span></span>

*   <span data-ttu-id="75a2d-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="75a2d-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-174">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-174">Requirements</span></span>

|<span data-ttu-id="75a2d-175">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-175">Requirement</span></span>| <span data-ttu-id="75a2d-176">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-177">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-178">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-178">1.0</span></span>|
|[<span data-ttu-id="75a2d-179">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-180">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-181">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-182">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-183">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="75a2d-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="75a2d-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="75a2d-185">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="75a2d-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="75a2d-p108">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="75a2d-p109">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-190">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-190">Type:</span></span>

*   <span data-ttu-id="75a2d-191">Chaîne</span><span class="sxs-lookup"><span data-stu-id="75a2d-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-192">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-192">Requirements</span></span>

|<span data-ttu-id="75a2d-193">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-193">Requirement</span></span>| <span data-ttu-id="75a2d-194">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-195">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-195">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-196">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-196">1.0</span></span>|
|[<span data-ttu-id="75a2d-197">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-198">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-200">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="75a2d-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="75a2d-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="75a2d-p110">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-204">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-204">Type:</span></span>

*   <span data-ttu-id="75a2d-205">Date</span><span class="sxs-lookup"><span data-stu-id="75a2d-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-206">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-206">Requirements</span></span>

|<span data-ttu-id="75a2d-207">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-207">Requirement</span></span>| <span data-ttu-id="75a2d-208">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-209">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-209">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-210">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-210">1.0</span></span>|
|[<span data-ttu-id="75a2d-211">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-212">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-213">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-214">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-215">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="75a2d-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="75a2d-216">dateTimeModified :Date</span></span>

<span data-ttu-id="75a2d-p111">Obtient la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-219">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="75a2d-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-220">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-220">Type:</span></span>

*   <span data-ttu-id="75a2d-221">Date</span><span class="sxs-lookup"><span data-stu-id="75a2d-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-222">Requirements</span></span>

|<span data-ttu-id="75a2d-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-223">Requirement</span></span>| <span data-ttu-id="75a2d-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-225">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-226">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-226">1.0</span></span>|
|[<span data-ttu-id="75a2d-227">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-228">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-229">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-230">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-231">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="75a2d-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="75a2d-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="75a2d-233">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="75a2d-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="75a2d-p112">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="75a2d-236">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-236">Read mode</span></span>

<span data-ttu-id="75a2d-237">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="75a2d-238">Mode composition</span><span class="sxs-lookup"><span data-stu-id="75a2d-238">Compose mode</span></span>

<span data-ttu-id="75a2d-239">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="75a2d-240">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="75a2d-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-241">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-241">Type:</span></span>

*   <span data-ttu-id="75a2d-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="75a2d-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-243">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-243">Requirements</span></span>

|<span data-ttu-id="75a2d-244">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-244">Requirement</span></span>| <span data-ttu-id="75a2d-245">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-246">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-246">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-247">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-247">1.0</span></span>|
|[<span data-ttu-id="75a2d-248">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-249">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-250">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-251">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-252">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-252">Example</span></span>

<span data-ttu-id="75a2d-253">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="75a2d-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="75a2d-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="75a2d-p113">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="75a2d-p114">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-259">Le `recipientType` propriété de la `EmailAddressDetails` objet dans le `from` propriété est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-260">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-260">Type:</span></span>

*   [<span data-ttu-id="75a2d-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="75a2d-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="75a2d-262">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-262">Requirements</span></span>

|<span data-ttu-id="75a2d-263">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-263">Requirement</span></span>| <span data-ttu-id="75a2d-264">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-265">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-266">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-266">1.0</span></span>|
|[<span data-ttu-id="75a2d-267">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-268">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-269">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-270">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="75a2d-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="75a2d-271">internetMessageId :String</span></span>

<span data-ttu-id="75a2d-p115">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-274">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-274">Type:</span></span>

*   <span data-ttu-id="75a2d-275">Chaîne</span><span class="sxs-lookup"><span data-stu-id="75a2d-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-276">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-276">Requirements</span></span>

|<span data-ttu-id="75a2d-277">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-277">Requirement</span></span>| <span data-ttu-id="75a2d-278">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-279">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-279">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-280">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-280">1.0</span></span>|
|[<span data-ttu-id="75a2d-281">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-282">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-283">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-284">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-285">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="75a2d-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="75a2d-286">itemClass :String</span></span>

<span data-ttu-id="75a2d-p116">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="75a2d-p117">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="75a2d-291">Type</span><span class="sxs-lookup"><span data-stu-id="75a2d-291">Type</span></span> | <span data-ttu-id="75a2d-292">Description</span><span class="sxs-lookup"><span data-stu-id="75a2d-292">Description</span></span> | <span data-ttu-id="75a2d-293">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="75a2d-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="75a2d-294">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="75a2d-294">Appointment items</span></span> | <span data-ttu-id="75a2d-295">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="75a2d-296">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="75a2d-296">Message items</span></span> | <span data-ttu-id="75a2d-297">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="75a2d-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="75a2d-298">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-299">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-299">Type:</span></span>

*   <span data-ttu-id="75a2d-300">Chaîne</span><span class="sxs-lookup"><span data-stu-id="75a2d-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-301">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-301">Requirements</span></span>

|<span data-ttu-id="75a2d-302">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-302">Requirement</span></span>| <span data-ttu-id="75a2d-303">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-304">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-304">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-305">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-305">1.0</span></span>|
|[<span data-ttu-id="75a2d-306">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-307">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-308">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-309">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-310">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="75a2d-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="75a2d-311">(nullable) itemId :String</span></span>

<span data-ttu-id="75a2d-p118">Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-314">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="75a2d-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="75a2d-315">Le `itemId` propriété n’est pas identique à l’identificateur d’entrée Outlook ou l’ID utilisé par l’API REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="75a2d-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="75a2d-316">Avant que les appels API REST à l’aide de cette valeur, elle doit être convertie à l’aide de `Office.context.mailbox.convertToRestId`, qui est disponible à compter d’exigence défini 1.3.</span><span class="sxs-lookup"><span data-stu-id="75a2d-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="75a2d-317">Pour plus d’informations, voir [utiliser les API REST d’Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="75a2d-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-318">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-318">Type:</span></span>

*   <span data-ttu-id="75a2d-319">Chaîne</span><span class="sxs-lookup"><span data-stu-id="75a2d-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-320">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-320">Requirements</span></span>

|<span data-ttu-id="75a2d-321">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-321">Requirement</span></span>| <span data-ttu-id="75a2d-322">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-323">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-323">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-324">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-324">1.0</span></span>|
|[<span data-ttu-id="75a2d-325">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-326">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-327">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-328">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-329">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-329">Example</span></span>

<span data-ttu-id="75a2d-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="75a2d-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="75a2d-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="75a2d-333">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="75a2d-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="75a2d-334">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="75a2d-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-335">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-335">Type:</span></span>

*   [<span data-ttu-id="75a2d-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="75a2d-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="75a2d-337">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-337">Requirements</span></span>

|<span data-ttu-id="75a2d-338">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-338">Requirement</span></span>| <span data-ttu-id="75a2d-339">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-340">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-341">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-341">1.0</span></span>|
|[<span data-ttu-id="75a2d-342">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-343">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-344">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-345">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-346">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="75a2d-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="75a2d-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="75a2d-348">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="75a2d-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="75a2d-349">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-349">Read mode</span></span>

<span data-ttu-id="75a2d-350">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="75a2d-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="75a2d-351">Mode composition</span><span class="sxs-lookup"><span data-stu-id="75a2d-351">Compose mode</span></span>

<span data-ttu-id="75a2d-352">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="75a2d-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-353">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-353">Type:</span></span>

*   <span data-ttu-id="75a2d-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="75a2d-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-355">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-355">Requirements</span></span>

|<span data-ttu-id="75a2d-356">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-356">Requirement</span></span>| <span data-ttu-id="75a2d-357">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-358">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-358">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-359">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-359">1.0</span></span>|
|[<span data-ttu-id="75a2d-360">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-361">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-362">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-363">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-364">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="75a2d-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="75a2d-365">normalizedSubject :String</span></span>

<span data-ttu-id="75a2d-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="75a2d-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject).</span><span class="sxs-lookup"><span data-stu-id="75a2d-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-370">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-370">Type:</span></span>

*   <span data-ttu-id="75a2d-371">Chaîne</span><span class="sxs-lookup"><span data-stu-id="75a2d-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-372">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-372">Requirements</span></span>

|<span data-ttu-id="75a2d-373">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-373">Requirement</span></span>| <span data-ttu-id="75a2d-374">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-375">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-375">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-376">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-376">1.0</span></span>|
|[<span data-ttu-id="75a2d-377">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-378">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-379">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-380">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-381">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="75a2d-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="75a2d-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="75a2d-383">Fournit l’accès à un événement les participants facultatifs.</span><span class="sxs-lookup"><span data-stu-id="75a2d-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="75a2d-384">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="75a2d-385">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-385">Read mode</span></span>

<span data-ttu-id="75a2d-386">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="75a2d-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="75a2d-387">Mode composition</span><span class="sxs-lookup"><span data-stu-id="75a2d-387">Compose mode</span></span>

<span data-ttu-id="75a2d-388">Le `optionalAttendees` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs à une réunion.</span><span class="sxs-lookup"><span data-stu-id="75a2d-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-389">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-389">Type:</span></span>

*   <span data-ttu-id="75a2d-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="75a2d-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-391">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-391">Requirements</span></span>

|<span data-ttu-id="75a2d-392">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-392">Requirement</span></span>| <span data-ttu-id="75a2d-393">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-394">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-394">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-395">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-395">1.0</span></span>|
|[<span data-ttu-id="75a2d-396">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-397">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-398">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-399">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-400">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="75a2d-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="75a2d-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="75a2d-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-404">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-404">Type:</span></span>

*   [<span data-ttu-id="75a2d-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="75a2d-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="75a2d-406">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-406">Requirements</span></span>

|<span data-ttu-id="75a2d-407">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-407">Requirement</span></span>| <span data-ttu-id="75a2d-408">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-409">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-410">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-410">1.0</span></span>|
|[<span data-ttu-id="75a2d-411">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-412">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-413">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-414">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-415">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="75a2d-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="75a2d-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="75a2d-417">Fournit l’accès à des participants obligatoires d’un événement.</span><span class="sxs-lookup"><span data-stu-id="75a2d-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="75a2d-418">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="75a2d-419">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-419">Read mode</span></span>

<span data-ttu-id="75a2d-420">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="75a2d-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="75a2d-421">Mode composition</span><span class="sxs-lookup"><span data-stu-id="75a2d-421">Compose mode</span></span>

<span data-ttu-id="75a2d-422">Le `requiredAttendees` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les participants obligatoires à une réunion.</span><span class="sxs-lookup"><span data-stu-id="75a2d-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-423">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-423">Type:</span></span>

*   <span data-ttu-id="75a2d-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="75a2d-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-425">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-425">Requirements</span></span>

|<span data-ttu-id="75a2d-426">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-426">Requirement</span></span>| <span data-ttu-id="75a2d-427">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-428">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-428">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-429">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-429">1.0</span></span>|
|[<span data-ttu-id="75a2d-430">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-431">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-432">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-433">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-434">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="75a2d-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="75a2d-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="75a2d-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="75a2d-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-440">Le `recipientType` propriété de la `EmailAddressDetails` objet dans le `from` propriété est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-440">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-441">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-441">Type:</span></span>

*   [<span data-ttu-id="75a2d-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="75a2d-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="75a2d-443">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-443">Requirements</span></span>

|<span data-ttu-id="75a2d-444">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-444">Requirement</span></span>| <span data-ttu-id="75a2d-445">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-446">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-446">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-447">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-447">1.0</span></span>|
|[<span data-ttu-id="75a2d-448">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-449">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-450">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-451">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-452">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="75a2d-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="75a2d-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="75a2d-454">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="75a2d-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="75a2d-p128">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="75a2d-457">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-457">Read mode</span></span>

<span data-ttu-id="75a2d-458">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="75a2d-459">Mode composition</span><span class="sxs-lookup"><span data-stu-id="75a2d-459">Compose mode</span></span>

<span data-ttu-id="75a2d-460">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="75a2d-461">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="75a2d-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-462">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-462">Type:</span></span>

*   <span data-ttu-id="75a2d-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="75a2d-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-464">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-464">Requirements</span></span>

|<span data-ttu-id="75a2d-465">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-465">Requirement</span></span>| <span data-ttu-id="75a2d-466">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-467">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-467">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-468">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-468">1.0</span></span>|
|[<span data-ttu-id="75a2d-469">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-470">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-471">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-472">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-473">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-473">Example</span></span>

<span data-ttu-id="75a2d-474">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="75a2d-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="75a2d-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="75a2d-476">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="75a2d-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="75a2d-477">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="75a2d-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="75a2d-478">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-478">Read mode</span></span>

<span data-ttu-id="75a2d-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="75a2d-481">Mode composition</span><span class="sxs-lookup"><span data-stu-id="75a2d-481">Compose mode</span></span>

<span data-ttu-id="75a2d-482">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="75a2d-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="75a2d-483">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-483">Type:</span></span>

*   <span data-ttu-id="75a2d-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="75a2d-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-485">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-485">Requirements</span></span>

|<span data-ttu-id="75a2d-486">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-486">Requirement</span></span>| <span data-ttu-id="75a2d-487">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-488">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-488">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-489">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-489">1.0</span></span>|
|[<span data-ttu-id="75a2d-490">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-491">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-492">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-493">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="75a2d-494">à : Array. <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[destinataires](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="75a2d-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="75a2d-495">Permet d’accéder aux destinataires de la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="75a2d-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="75a2d-496">Le type d’objet et le niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="75a2d-497">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-497">Read mode</span></span>

<span data-ttu-id="75a2d-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="75a2d-500">Mode composition</span><span class="sxs-lookup"><span data-stu-id="75a2d-500">Compose mode</span></span>

<span data-ttu-id="75a2d-501">Le `to` propriété retourne un `Recipients` objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="75a2d-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="75a2d-502">Type :</span><span class="sxs-lookup"><span data-stu-id="75a2d-502">Type:</span></span>

*   <span data-ttu-id="75a2d-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="75a2d-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-504">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-504">Requirements</span></span>

|<span data-ttu-id="75a2d-505">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-505">Requirement</span></span>| <span data-ttu-id="75a2d-506">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-507">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-508">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-508">1.0</span></span>|
|[<span data-ttu-id="75a2d-509">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-510">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-511">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-512">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-513">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="75a2d-514">Méthodes</span><span class="sxs-lookup"><span data-stu-id="75a2d-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="75a2d-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="75a2d-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="75a2d-516">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="75a2d-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="75a2d-517">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="75a2d-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="75a2d-518">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="75a2d-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="75a2d-519">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="75a2d-519">Parameters:</span></span>

|<span data-ttu-id="75a2d-520">Nom</span><span class="sxs-lookup"><span data-stu-id="75a2d-520">Name</span></span>| <span data-ttu-id="75a2d-521">Type</span><span class="sxs-lookup"><span data-stu-id="75a2d-521">Type</span></span>| <span data-ttu-id="75a2d-522">Attributs</span><span class="sxs-lookup"><span data-stu-id="75a2d-522">Attributes</span></span>| <span data-ttu-id="75a2d-523">Description</span><span class="sxs-lookup"><span data-stu-id="75a2d-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="75a2d-524">String</span><span class="sxs-lookup"><span data-stu-id="75a2d-524">String</span></span>||<span data-ttu-id="75a2d-p132">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="75a2d-527">String</span><span class="sxs-lookup"><span data-stu-id="75a2d-527">String</span></span>||<span data-ttu-id="75a2d-p133">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="75a2d-530">Objet</span><span class="sxs-lookup"><span data-stu-id="75a2d-530">Object</span></span>| <span data-ttu-id="75a2d-531">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-531">&lt;optional&gt;</span></span>|<span data-ttu-id="75a2d-532">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="75a2d-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="75a2d-533">Objet</span><span class="sxs-lookup"><span data-stu-id="75a2d-533">Object</span></span>| <span data-ttu-id="75a2d-534">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-534">&lt;optional&gt;</span></span>|<span data-ttu-id="75a2d-535">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="75a2d-536">fonction</span><span class="sxs-lookup"><span data-stu-id="75a2d-536">function</span></span>| <span data-ttu-id="75a2d-537">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-537">&lt;optional&gt;</span></span>|<span data-ttu-id="75a2d-538">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="75a2d-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="75a2d-539">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="75a2d-540">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="75a2d-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="75a2d-541">Erreurs</span><span class="sxs-lookup"><span data-stu-id="75a2d-541">Errors</span></span>

| <span data-ttu-id="75a2d-542">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="75a2d-542">Error code</span></span> | <span data-ttu-id="75a2d-543">Description</span><span class="sxs-lookup"><span data-stu-id="75a2d-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="75a2d-544">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="75a2d-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="75a2d-545">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="75a2d-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="75a2d-546">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="75a2d-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="75a2d-547">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-547">Requirements</span></span>

|<span data-ttu-id="75a2d-548">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-548">Requirement</span></span>| <span data-ttu-id="75a2d-549">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-550">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-550">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-551">1.1</span><span class="sxs-lookup"><span data-stu-id="75a2d-551">1.1</span></span>|
|[<span data-ttu-id="75a2d-552">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="75a2d-554">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-555">Composition</span><span class="sxs-lookup"><span data-stu-id="75a2d-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-556">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-556">Example</span></span>

```JavaScript
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="75a2d-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="75a2d-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="75a2d-558">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="75a2d-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="75a2d-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="75a2d-562">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="75a2d-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="75a2d-563">Si votre complément Office est en cours d’exécution dans Outlook Web App, le `addItemAttachmentAsync` méthode permettre joindre des éléments à des éléments autres que l’élément à modifier ; Toutefois, cela n’est pas pris en charge et n’est pas recommandée.</span><span class="sxs-lookup"><span data-stu-id="75a2d-563">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="75a2d-564">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="75a2d-564">Parameters:</span></span>

|<span data-ttu-id="75a2d-565">Nom</span><span class="sxs-lookup"><span data-stu-id="75a2d-565">Name</span></span>| <span data-ttu-id="75a2d-566">Type</span><span class="sxs-lookup"><span data-stu-id="75a2d-566">Type</span></span>| <span data-ttu-id="75a2d-567">Attributs</span><span class="sxs-lookup"><span data-stu-id="75a2d-567">Attributes</span></span>| <span data-ttu-id="75a2d-568">Description</span><span class="sxs-lookup"><span data-stu-id="75a2d-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="75a2d-569">String</span><span class="sxs-lookup"><span data-stu-id="75a2d-569">String</span></span>||<span data-ttu-id="75a2d-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="75a2d-572">String</span><span class="sxs-lookup"><span data-stu-id="75a2d-572">String</span></span>||<span data-ttu-id="75a2d-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="75a2d-575">Objet</span><span class="sxs-lookup"><span data-stu-id="75a2d-575">Object</span></span>| <span data-ttu-id="75a2d-576">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-576">&lt;optional&gt;</span></span>|<span data-ttu-id="75a2d-577">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="75a2d-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="75a2d-578">Objet</span><span class="sxs-lookup"><span data-stu-id="75a2d-578">Object</span></span>| <span data-ttu-id="75a2d-579">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-579">&lt;optional&gt;</span></span>|<span data-ttu-id="75a2d-580">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="75a2d-581">fonction</span><span class="sxs-lookup"><span data-stu-id="75a2d-581">function</span></span>| <span data-ttu-id="75a2d-582">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-582">&lt;optional&gt;</span></span>|<span data-ttu-id="75a2d-583">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="75a2d-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="75a2d-584">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="75a2d-585">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="75a2d-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="75a2d-586">Erreurs</span><span class="sxs-lookup"><span data-stu-id="75a2d-586">Errors</span></span>

| <span data-ttu-id="75a2d-587">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="75a2d-587">Error code</span></span> | <span data-ttu-id="75a2d-588">Description</span><span class="sxs-lookup"><span data-stu-id="75a2d-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="75a2d-589">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="75a2d-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="75a2d-590">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-590">Requirements</span></span>

|<span data-ttu-id="75a2d-591">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-591">Requirement</span></span>| <span data-ttu-id="75a2d-592">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-593">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-593">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-594">1.1</span><span class="sxs-lookup"><span data-stu-id="75a2d-594">1.1</span></span>|
|[<span data-ttu-id="75a2d-595">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="75a2d-597">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-598">Composition</span><span class="sxs-lookup"><span data-stu-id="75a2d-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-599">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-599">Example</span></span>

<span data-ttu-id="75a2d-600">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="75a2d-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="75a2d-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="75a2d-602">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="75a2d-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-603">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="75a2d-603">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="75a2d-604">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="75a2d-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="75a2d-605">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="75a2d-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-606">La possibilité d’inclure des pièces jointes dans l’appel à `displayReplyAllForm` n’est pas pris en charge dans un ensemble de conditions requises 1.1.</span><span class="sxs-lookup"><span data-stu-id="75a2d-606">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="75a2d-607">La prise en charge des pièces jointes a été ajoutée à `displayReplyAllForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="75a2d-607">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="75a2d-608">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="75a2d-608">Parameters:</span></span>

|<span data-ttu-id="75a2d-609">Nom</span><span class="sxs-lookup"><span data-stu-id="75a2d-609">Name</span></span>| <span data-ttu-id="75a2d-610">Type</span><span class="sxs-lookup"><span data-stu-id="75a2d-610">Type</span></span>| <span data-ttu-id="75a2d-611">Description</span><span class="sxs-lookup"><span data-stu-id="75a2d-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="75a2d-612">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="75a2d-612">String &#124; Object</span></span>| |<span data-ttu-id="75a2d-p138">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="75a2d-615">**OU**</span><span class="sxs-lookup"><span data-stu-id="75a2d-615">**OR**</span></span><br/><span data-ttu-id="75a2d-p139">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="75a2d-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="75a2d-618">String</span><span class="sxs-lookup"><span data-stu-id="75a2d-618">String</span></span> | <span data-ttu-id="75a2d-619">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-619">&lt;optional&gt;</span></span> | <span data-ttu-id="75a2d-p140">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="75a2d-622">function</span><span class="sxs-lookup"><span data-stu-id="75a2d-622">function</span></span> | <span data-ttu-id="75a2d-623">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-623">&lt;optional&gt;</span></span> | <span data-ttu-id="75a2d-624">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="75a2d-624">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="75a2d-625">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-625">Requirements</span></span>

|<span data-ttu-id="75a2d-626">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-626">Requirement</span></span>| <span data-ttu-id="75a2d-627">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-628">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-628">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-629">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-629">1.0</span></span>|
|[<span data-ttu-id="75a2d-630">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-631">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-632">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-633">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-633">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="75a2d-634">Exemples</span><span class="sxs-lookup"><span data-stu-id="75a2d-634">Examples</span></span>

<span data-ttu-id="75a2d-635">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-635">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="75a2d-636">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="75a2d-636">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="75a2d-637">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="75a2d-637">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="75a2d-638">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-638">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="75a2d-639">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="75a2d-639">displayReplyForm(formData)</span></span>

<span data-ttu-id="75a2d-640">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="75a2d-640">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-641">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="75a2d-641">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="75a2d-642">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="75a2d-642">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="75a2d-643">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="75a2d-643">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-644">La possibilité d’inclure des pièces jointes dans l’appel à `displayReplyForm` n’est pas pris en charge dans un ensemble de conditions requises 1.1.</span><span class="sxs-lookup"><span data-stu-id="75a2d-644">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="75a2d-645">La prise en charge des pièces jointes a été ajoutée à `displayReplyForm` dans les versions d’ensemble de conditions requises 1.2 et supérieures.</span><span class="sxs-lookup"><span data-stu-id="75a2d-645">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="75a2d-646">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="75a2d-646">Parameters:</span></span>

|<span data-ttu-id="75a2d-647">Nom</span><span class="sxs-lookup"><span data-stu-id="75a2d-647">Name</span></span>| <span data-ttu-id="75a2d-648">Type</span><span class="sxs-lookup"><span data-stu-id="75a2d-648">Type</span></span>| <span data-ttu-id="75a2d-649">Description</span><span class="sxs-lookup"><span data-stu-id="75a2d-649">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="75a2d-650">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="75a2d-650">String &#124; Object</span></span>| | <span data-ttu-id="75a2d-p142">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="75a2d-653">**OU**</span><span class="sxs-lookup"><span data-stu-id="75a2d-653">**OR**</span></span><br/><span data-ttu-id="75a2d-p143">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="75a2d-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="75a2d-656">String</span><span class="sxs-lookup"><span data-stu-id="75a2d-656">String</span></span> | <span data-ttu-id="75a2d-657">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-657">&lt;optional&gt;</span></span> | <span data-ttu-id="75a2d-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="75a2d-660">function</span><span class="sxs-lookup"><span data-stu-id="75a2d-660">function</span></span> | <span data-ttu-id="75a2d-661">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-661">&lt;optional&gt;</span></span> | <span data-ttu-id="75a2d-662">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="75a2d-662">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="75a2d-663">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-663">Requirements</span></span>

|<span data-ttu-id="75a2d-664">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-664">Requirement</span></span>| <span data-ttu-id="75a2d-665">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-665">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-666">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-666">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-667">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-667">1.0</span></span>|
|[<span data-ttu-id="75a2d-668">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-668">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-669">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-669">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-670">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-670">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-671">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-671">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="75a2d-672">Exemples</span><span class="sxs-lookup"><span data-stu-id="75a2d-672">Examples</span></span>

<span data-ttu-id="75a2d-673">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-673">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="75a2d-674">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="75a2d-674">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="75a2d-675">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="75a2d-675">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="75a2d-676">Réponse avec un corps et un rappel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-676">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="75a2d-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="75a2d-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="75a2d-678">Obtient les entités trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="75a2d-678">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-679">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="75a2d-679">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-680">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-680">Requirements</span></span>

|<span data-ttu-id="75a2d-681">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-681">Requirement</span></span>| <span data-ttu-id="75a2d-682">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-682">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-683">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-683">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-684">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-684">1.0</span></span>|
|[<span data-ttu-id="75a2d-685">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-685">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-686">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-687">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-687">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-688">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-688">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="75a2d-689">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="75a2d-689">Returns:</span></span>

<span data-ttu-id="75a2d-690">Type : [Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="75a2d-690">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="75a2d-691">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-691">Example</span></span>

<span data-ttu-id="75a2d-692">L’exemple suivant accède aux entités de contacts dans le corps de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-692">The following example accesses the contacts entities in the current item's body.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="75a2d-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="75a2d-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="75a2d-694">Obtient un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="75a2d-694">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-695">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="75a2d-695">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="75a2d-696">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="75a2d-696">Parameters:</span></span>

|<span data-ttu-id="75a2d-697">Nom</span><span class="sxs-lookup"><span data-stu-id="75a2d-697">Name</span></span>| <span data-ttu-id="75a2d-698">Type</span><span class="sxs-lookup"><span data-stu-id="75a2d-698">Type</span></span>| <span data-ttu-id="75a2d-699">Description</span><span class="sxs-lookup"><span data-stu-id="75a2d-699">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="75a2d-700">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="75a2d-700">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="75a2d-701">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="75a2d-701">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="75a2d-702">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-702">Requirements</span></span>

|<span data-ttu-id="75a2d-703">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-703">Requirement</span></span>| <span data-ttu-id="75a2d-704">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-704">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-705">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-705">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-706">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-706">1.0</span></span>|
|[<span data-ttu-id="75a2d-707">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-707">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-708">Restricted</span><span class="sxs-lookup"><span data-stu-id="75a2d-708">Restricted</span></span>|
|[<span data-ttu-id="75a2d-709">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-709">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-710">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-710">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="75a2d-711">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="75a2d-711">Returns:</span></span>

<span data-ttu-id="75a2d-712">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="75a2d-712">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="75a2d-713">Si aucune entité du type spécifié n’est présentes dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="75a2d-713">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="75a2d-714">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-714">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="75a2d-715">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="75a2d-715">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="75a2d-716">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="75a2d-716">Value of `entityType`</span></span> | <span data-ttu-id="75a2d-717">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="75a2d-717">Type of objects in returned array</span></span> | <span data-ttu-id="75a2d-718">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="75a2d-718">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="75a2d-719">String</span><span class="sxs-lookup"><span data-stu-id="75a2d-719">String</span></span> | <span data-ttu-id="75a2d-720">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="75a2d-720">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="75a2d-721">Contact</span><span class="sxs-lookup"><span data-stu-id="75a2d-721">Contact</span></span> | <span data-ttu-id="75a2d-722">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="75a2d-722">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="75a2d-723">String</span><span class="sxs-lookup"><span data-stu-id="75a2d-723">String</span></span> | <span data-ttu-id="75a2d-724">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="75a2d-724">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="75a2d-725">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="75a2d-725">MeetingSuggestion</span></span> | <span data-ttu-id="75a2d-726">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="75a2d-726">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="75a2d-727">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="75a2d-727">PhoneNumber</span></span> | <span data-ttu-id="75a2d-728">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="75a2d-728">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="75a2d-729">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="75a2d-729">TaskSuggestion</span></span> | <span data-ttu-id="75a2d-730">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="75a2d-730">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="75a2d-731">String</span><span class="sxs-lookup"><span data-stu-id="75a2d-731">String</span></span> | <span data-ttu-id="75a2d-732">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="75a2d-732">**Restricted**</span></span> |

<span data-ttu-id="75a2d-733">Type :  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="75a2d-733">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="75a2d-734">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-734">Example</span></span>

<span data-ttu-id="75a2d-735">L’exemple suivant montre comment accéder à un tableau de chaînes qui représentent les adresses postales dans le corps de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-735">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```JavaScript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="75a2d-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="75a2d-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="75a2d-737">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="75a2d-737">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-738">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="75a2d-738">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="75a2d-739">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="75a2d-739">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="75a2d-740">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="75a2d-740">Parameters:</span></span>

|<span data-ttu-id="75a2d-741">Nom</span><span class="sxs-lookup"><span data-stu-id="75a2d-741">Name</span></span>| <span data-ttu-id="75a2d-742">Type</span><span class="sxs-lookup"><span data-stu-id="75a2d-742">Type</span></span>| <span data-ttu-id="75a2d-743">Description</span><span class="sxs-lookup"><span data-stu-id="75a2d-743">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="75a2d-744">String</span><span class="sxs-lookup"><span data-stu-id="75a2d-744">String</span></span>|<span data-ttu-id="75a2d-745">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="75a2d-745">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="75a2d-746">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-746">Requirements</span></span>

|<span data-ttu-id="75a2d-747">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-747">Requirement</span></span>| <span data-ttu-id="75a2d-748">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-748">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-749">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-749">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-750">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-750">1.0</span></span>|
|[<span data-ttu-id="75a2d-751">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-751">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-752">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-752">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-753">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-753">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-754">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-754">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="75a2d-755">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="75a2d-755">Returns:</span></span>

<span data-ttu-id="75a2d-p146">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="75a2d-758">Type : Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="75a2d-758">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="75a2d-759">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="75a2d-759">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="75a2d-760">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="75a2d-760">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-761">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="75a2d-761">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="75a2d-p147">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="75a2d-765">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="75a2d-765">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="75a2d-766">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-766">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="75a2d-p148">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="75a2d-769">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-769">Requirements</span></span>

|<span data-ttu-id="75a2d-770">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-770">Requirement</span></span>| <span data-ttu-id="75a2d-771">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-772">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-772">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-773">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-773">1.0</span></span>|
|[<span data-ttu-id="75a2d-774">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-774">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-775">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-776">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-776">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-777">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-777">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="75a2d-778">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="75a2d-778">Returns:</span></span>

<span data-ttu-id="75a2d-p149">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="75a2d-781">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="75a2d-781">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="75a2d-782">Object</span><span class="sxs-lookup"><span data-stu-id="75a2d-782">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="75a2d-783">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-783">Example</span></span>

<span data-ttu-id="75a2d-784">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments<rule> d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</rule></span><span class="sxs-lookup"><span data-stu-id="75a2d-784">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="75a2d-785">getregexmatchesbyname (Name) → (nullable) {tableau. < chaîne >}</span><span class="sxs-lookup"><span data-stu-id="75a2d-785">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="75a2d-786">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="75a2d-786">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="75a2d-787">Cette méthode n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="75a2d-787">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="75a2d-788">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="75a2d-788">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="75a2d-p150">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="75a2d-791">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="75a2d-791">Parameters:</span></span>

|<span data-ttu-id="75a2d-792">Nom</span><span class="sxs-lookup"><span data-stu-id="75a2d-792">Name</span></span>| <span data-ttu-id="75a2d-793">Type</span><span class="sxs-lookup"><span data-stu-id="75a2d-793">Type</span></span>| <span data-ttu-id="75a2d-794">Description</span><span class="sxs-lookup"><span data-stu-id="75a2d-794">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="75a2d-795">String</span><span class="sxs-lookup"><span data-stu-id="75a2d-795">String</span></span>|<span data-ttu-id="75a2d-796">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="75a2d-796">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="75a2d-797">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-797">Requirements</span></span>

|<span data-ttu-id="75a2d-798">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-798">Requirement</span></span>| <span data-ttu-id="75a2d-799">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-799">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-800">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-800">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-801">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-801">1.0</span></span>|
|[<span data-ttu-id="75a2d-802">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-802">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-803">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-803">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-804">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-804">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-805">Lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-805">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="75a2d-806">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="75a2d-806">Returns:</span></span>

<span data-ttu-id="75a2d-807">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="75a2d-807">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="75a2d-808">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="75a2d-808">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="75a2d-809">< Chaîne > tableau.</span><span class="sxs-lookup"><span data-stu-id="75a2d-809">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="75a2d-810">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-810">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="75a2d-811">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="75a2d-811">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="75a2d-812">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="75a2d-812">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="75a2d-p151">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="75a2d-816">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="75a2d-816">Parameters:</span></span>

|<span data-ttu-id="75a2d-817">Nom</span><span class="sxs-lookup"><span data-stu-id="75a2d-817">Name</span></span>| <span data-ttu-id="75a2d-818">Type</span><span class="sxs-lookup"><span data-stu-id="75a2d-818">Type</span></span>| <span data-ttu-id="75a2d-819">Attributs</span><span class="sxs-lookup"><span data-stu-id="75a2d-819">Attributes</span></span>| <span data-ttu-id="75a2d-820">Description</span><span class="sxs-lookup"><span data-stu-id="75a2d-820">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="75a2d-821">function</span><span class="sxs-lookup"><span data-stu-id="75a2d-821">function</span></span>||<span data-ttu-id="75a2d-822">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="75a2d-822">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="75a2d-823">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="75a2d-823">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="75a2d-824">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées de l’élément et enregistrer les modifications à la propriété personnalisée est redéfinie sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="75a2d-824">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="75a2d-825">Objet</span><span class="sxs-lookup"><span data-stu-id="75a2d-825">Object</span></span>| <span data-ttu-id="75a2d-826">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-826">&lt;optional&gt;</span></span>|<span data-ttu-id="75a2d-827">Les développeurs peuvent fournir n’importe quel objet qu’ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-827">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="75a2d-828">Cet objet est accessible par le `asyncResult.asyncContext` propriété dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-828">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="75a2d-829">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-829">Requirements</span></span>

|<span data-ttu-id="75a2d-830">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-830">Requirement</span></span>| <span data-ttu-id="75a2d-831">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-831">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-832">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-832">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-833">1.0</span><span class="sxs-lookup"><span data-stu-id="75a2d-833">1.0</span></span>|
|[<span data-ttu-id="75a2d-834">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-834">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-835">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-835">ReadItem</span></span>|
|[<span data-ttu-id="75a2d-836">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-836">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-837">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="75a2d-837">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-838">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-838">Example</span></span>

<span data-ttu-id="75a2d-p154">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="75a2d-842">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="75a2d-842">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="75a2d-843">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="75a2d-843">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="75a2d-p155">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="75a2d-848">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="75a2d-848">Parameters:</span></span>

|<span data-ttu-id="75a2d-849">Nom</span><span class="sxs-lookup"><span data-stu-id="75a2d-849">Name</span></span>| <span data-ttu-id="75a2d-850">Type</span><span class="sxs-lookup"><span data-stu-id="75a2d-850">Type</span></span>| <span data-ttu-id="75a2d-851">Attributs</span><span class="sxs-lookup"><span data-stu-id="75a2d-851">Attributes</span></span>| <span data-ttu-id="75a2d-852">Description</span><span class="sxs-lookup"><span data-stu-id="75a2d-852">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="75a2d-853">String</span><span class="sxs-lookup"><span data-stu-id="75a2d-853">String</span></span>||<span data-ttu-id="75a2d-p156">Identificateur de la pièce jointe à supprimer. La longueur maximale de la chaîne est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="75a2d-p156">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="75a2d-856">Objet</span><span class="sxs-lookup"><span data-stu-id="75a2d-856">Object</span></span>| <span data-ttu-id="75a2d-857">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-857">&lt;optional&gt;</span></span>|<span data-ttu-id="75a2d-858">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="75a2d-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="75a2d-859">Objet</span><span class="sxs-lookup"><span data-stu-id="75a2d-859">Object</span></span>| <span data-ttu-id="75a2d-860">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-860">&lt;optional&gt;</span></span>|<span data-ttu-id="75a2d-861">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="75a2d-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="75a2d-862">fonction</span><span class="sxs-lookup"><span data-stu-id="75a2d-862">function</span></span>| <span data-ttu-id="75a2d-863">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="75a2d-863">&lt;optional&gt;</span></span>|<span data-ttu-id="75a2d-864">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="75a2d-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="75a2d-865">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="75a2d-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="75a2d-866">Erreurs</span><span class="sxs-lookup"><span data-stu-id="75a2d-866">Errors</span></span>

| <span data-ttu-id="75a2d-867">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="75a2d-867">Error code</span></span> | <span data-ttu-id="75a2d-868">Description</span><span class="sxs-lookup"><span data-stu-id="75a2d-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="75a2d-869">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="75a2d-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="75a2d-870">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="75a2d-870">Requirements</span></span>

|<span data-ttu-id="75a2d-871">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="75a2d-871">Requirement</span></span>| <span data-ttu-id="75a2d-872">Valeur</span><span class="sxs-lookup"><span data-stu-id="75a2d-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="75a2d-873">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="75a2d-873">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75a2d-874">1.1</span><span class="sxs-lookup"><span data-stu-id="75a2d-874">1.1</span></span>|
|[<span data-ttu-id="75a2d-875">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="75a2d-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75a2d-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="75a2d-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="75a2d-877">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="75a2d-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="75a2d-878">Composition</span><span class="sxs-lookup"><span data-stu-id="75a2d-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="75a2d-879">Exemple</span><span class="sxs-lookup"><span data-stu-id="75a2d-879">Example</span></span>

<span data-ttu-id="75a2d-880">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="75a2d-880">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```