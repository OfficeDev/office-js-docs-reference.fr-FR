
# <a name="context"></a><span data-ttu-id="39f35-101">context</span><span class="sxs-lookup"><span data-stu-id="39f35-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="39f35-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="39f35-102">[Office](Office.md).context</span></span>

<span data-ttu-id="39f35-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API partagée](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="39f35-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f35-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="39f35-105">Requirements</span></span>

|<span data-ttu-id="39f35-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="39f35-106">Requirement</span></span>| <span data-ttu-id="39f35-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="39f35-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f35-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="39f35-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f35-109">1.0</span><span class="sxs-lookup"><span data-stu-id="39f35-109">1.0</span></span>|
|[<span data-ttu-id="39f35-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="39f35-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39f35-111">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="39f35-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="39f35-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="39f35-112">Members and methods</span></span>

| <span data-ttu-id="39f35-113">Membre</span><span class="sxs-lookup"><span data-stu-id="39f35-113">Member</span></span> | <span data-ttu-id="39f35-114">Type</span><span class="sxs-lookup"><span data-stu-id="39f35-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="39f35-115">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="39f35-115">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="39f35-116">Membre</span><span class="sxs-lookup"><span data-stu-id="39f35-116">Member</span></span> |
| [<span data-ttu-id="39f35-117">officeTheme</span><span class="sxs-lookup"><span data-stu-id="39f35-117">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="39f35-118">Membre</span><span class="sxs-lookup"><span data-stu-id="39f35-118">Member</span></span> |
| [<span data-ttu-id="39f35-119">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="39f35-119">roamingSettings</span></span>](#roamingsettings-roamingsettingsjavascriptapioutlook16officeroamingsettings) | <span data-ttu-id="39f35-120">Membre</span><span class="sxs-lookup"><span data-stu-id="39f35-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="39f35-121">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="39f35-121">Namespaces</span></span>

<span data-ttu-id="39f35-122">[boîte aux lettres](office.context.mailbox.md): permet d’accéder au modèle objet de complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="39f35-122">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="39f35-123">Membres</span><span class="sxs-lookup"><span data-stu-id="39f35-123">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="39f35-124">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="39f35-124">displayLanguage :String</span></span>

<span data-ttu-id="39f35-125">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="39f35-125">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="39f35-126">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="39f35-126">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="39f35-127">Type :</span><span class="sxs-lookup"><span data-stu-id="39f35-127">Type:</span></span>

*   <span data-ttu-id="39f35-128">Chaîne</span><span class="sxs-lookup"><span data-stu-id="39f35-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="39f35-129">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="39f35-129">Requirements</span></span>

|<span data-ttu-id="39f35-130">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="39f35-130">Requirement</span></span>| <span data-ttu-id="39f35-131">Valeur</span><span class="sxs-lookup"><span data-stu-id="39f35-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f35-132">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="39f35-132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f35-133">1.0</span><span class="sxs-lookup"><span data-stu-id="39f35-133">1.0</span></span>|
|[<span data-ttu-id="39f35-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="39f35-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39f35-135">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="39f35-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f35-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="39f35-136">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="officetheme-object"></a><span data-ttu-id="39f35-137">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="39f35-137">officeTheme :Object</span></span>

<span data-ttu-id="39f35-138">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="39f35-138">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="39f35-139">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="39f35-139">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="39f35-p102">À l’aide des couleurs du thème Office, vous pouvez coordonner le modèle de couleurs de votre complément avec le thème Office actuel sélectionné par l’utilisateur dans **Fichier > Compte Office > Thème Office**, qui est appliqué à toutes les applications hôtes Office. Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="39f35-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="39f35-142">Type :</span><span class="sxs-lookup"><span data-stu-id="39f35-142">Type:</span></span>

*   <span data-ttu-id="39f35-143">Objet</span><span class="sxs-lookup"><span data-stu-id="39f35-143">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="39f35-144">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="39f35-144">Properties:</span></span>

|<span data-ttu-id="39f35-145">Nom</span><span class="sxs-lookup"><span data-stu-id="39f35-145">Name</span></span>| <span data-ttu-id="39f35-146">Type</span><span class="sxs-lookup"><span data-stu-id="39f35-146">Type</span></span>| <span data-ttu-id="39f35-147">Description</span><span class="sxs-lookup"><span data-stu-id="39f35-147">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="39f35-148">String</span><span class="sxs-lookup"><span data-stu-id="39f35-148">String</span></span>|<span data-ttu-id="39f35-149">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="39f35-149">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="39f35-150">String</span><span class="sxs-lookup"><span data-stu-id="39f35-150">String</span></span>|<span data-ttu-id="39f35-151">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="39f35-151">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="39f35-152">String</span><span class="sxs-lookup"><span data-stu-id="39f35-152">String</span></span>|<span data-ttu-id="39f35-153">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="39f35-153">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="39f35-154">String</span><span class="sxs-lookup"><span data-stu-id="39f35-154">String</span></span>|<span data-ttu-id="39f35-155">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="39f35-155">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="39f35-156">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="39f35-156">Requirements</span></span>

|<span data-ttu-id="39f35-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="39f35-157">Requirement</span></span>| <span data-ttu-id="39f35-158">Valeur</span><span class="sxs-lookup"><span data-stu-id="39f35-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f35-159">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="39f35-159">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f35-160">1.3</span><span class="sxs-lookup"><span data-stu-id="39f35-160">1.3</span></span>|
|[<span data-ttu-id="39f35-161">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="39f35-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39f35-162">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="39f35-162">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="39f35-163">Exemple</span><span class="sxs-lookup"><span data-stu-id="39f35-163">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook16officeroamingsettings"></a><span data-ttu-id="39f35-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_6/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="39f35-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_6/office.RoamingSettings)</span></span>

<span data-ttu-id="39f35-165">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="39f35-165">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="39f35-166">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="39f35-166">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="39f35-167">Type :</span><span class="sxs-lookup"><span data-stu-id="39f35-167">Type:</span></span>

*   [<span data-ttu-id="39f35-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="39f35-168">RoamingSettings</span></span>](/javascript/api/outlook_1_6/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="39f35-169">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="39f35-169">Requirements</span></span>

|<span data-ttu-id="39f35-170">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="39f35-170">Requirement</span></span>| <span data-ttu-id="39f35-171">Valeur</span><span class="sxs-lookup"><span data-stu-id="39f35-171">Value</span></span>|
|---|---|
|[<span data-ttu-id="39f35-172">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="39f35-172">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39f35-173">1.0</span><span class="sxs-lookup"><span data-stu-id="39f35-173">1.0</span></span>|
|[<span data-ttu-id="39f35-174">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="39f35-174">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39f35-175">Restricted</span><span class="sxs-lookup"><span data-stu-id="39f35-175">Restricted</span></span>|
|[<span data-ttu-id="39f35-176">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="39f35-176">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="39f35-177">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="39f35-177">Compose or read</span></span>|