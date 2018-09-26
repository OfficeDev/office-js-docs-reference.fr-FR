# <a name="word-javascript-api-overview"></a><span data-ttu-id="d49ab-101">Vue d’ensemble de l’API JavaScript de Word</span><span class="sxs-lookup"><span data-stu-id="d49ab-101">Word JavaScript API overview</span></span>

<span data-ttu-id="d49ab-p101">Word propose un ensemble complet d’API que vous pouvez utiliser pour créer des compléments qui interagissent avec les métadonnées et le contenu du document. Ces API permettent de créer des expériences attrayantes qui s’intègrent à Word et l’étendent. Vous pouvez importer et exporter du contenu, assembler de nouveaux documents provenant de différentes sources de données et réaliser une intégration avec des flux de travail de document pour créer des solutions de document personnalisées.</span><span class="sxs-lookup"><span data-stu-id="d49ab-p101">Word provides a rich set of APIs that you can use to create add-ins that interact with document content and metadata. Use these APIs to create compelling experiences that integrate with and extend Word. You can import and export content, assemble new documents from different data sources, and integrate with document workflows to create custom document solutions.</span></span>

<span data-ttu-id="d49ab-105">Vous pouvez utiliser deux API JavaScript pour interagir avec les objets et les métadonnées d’un document Word :</span><span class="sxs-lookup"><span data-stu-id="d49ab-105">You can use two JavaScript APIs to interact with the objects and metadata in a Word document:</span></span>

- <span data-ttu-id="d49ab-106">API JavaScript pour Word : introduite dans Office 2016.</span><span class="sxs-lookup"><span data-stu-id="d49ab-106">Word JavaScript API - Introduced in Office 2016.</span></span>
- <span data-ttu-id="d49ab-107">[Interface API JavaScript pour Office](../javascript-api-for-office.md) (Office.js) : introduite dans Office 2013.</span><span class="sxs-lookup"><span data-stu-id="d49ab-107">[JavaScript API for Office](../javascript-api-for-office.md) (Office.js) - Introduced in Office 2013.</span></span>

## <a name="word-javascript-api"></a><span data-ttu-id="d49ab-108">API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="d49ab-108">Word JavaScript API</span></span>

<span data-ttu-id="d49ab-p102">L’API JavaScript pour Word est chargée par Office.js. Elle offre une nouvelle façon d’interagir avec les objets tels que les documents et les paragraphes. Ainsi, vous n’utilisez plus d’API asynchrones individuelles pour extraire et mettre à jour chacun de ces objets. L’API JavaScript pour Word fournit des objets JavaScript de « proxy » qui correspondent aux objets réels utilisés dans Word. Vous pouvez interagir avec ces objets de proxy en lisant et en écrivant leurs propriétés de façon synchronisée, et en appelant des méthodes synchrones pour effectuer des opérations les concernant. Ces interactions avec les objets de proxy ne sont pas immédiatement appliquées dans le script en cours d’exécution. La méthode **context.sync** synchronise l’état de vos objets JavaScript en cours d’exécution et celui des objets réels en exécutant des instructions mises en file d’attente et en récupérant les propriétés des objets Word chargés pour les utiliser dans votre script.</span><span class="sxs-lookup"><span data-stu-id="d49ab-p102">The Word JavaScript API is loaded by Office.js. The Word JavaScript API changes the way that you can interact with objects like documents and paragraphs. Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the Word JavaScript API provides “proxy” JavaScript objects that correspond to the real objects running in Word. You can interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them. These interactions with proxy objects aren’t immediately realized in the running script. The **context.sync** method synchronizes the state between your running JavaScript and the real objects in Office by executing queued instructions and retrieving properties of loaded Word objects for use in your script.</span></span>

## <a name="javascript-api-for-office"></a><span data-ttu-id="d49ab-115">Interface API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="d49ab-115">JavaScript API for Office</span></span>

<span data-ttu-id="d49ab-116">Vous pouvez référencer Office.js à partir des emplacements suivants :</span><span class="sxs-lookup"><span data-stu-id="d49ab-116">You can reference Office.js from the following locations:</span></span>

* <span data-ttu-id="d49ab-117">https://appsforoffice.microsoft.com/lib/1/hosted/office.js-Utilisez cette ressource pour les compléments de production.</span><span class="sxs-lookup"><span data-stu-id="d49ab-117">https://appsforoffice.microsoft.com/lib/1/hosted/office.js - use this resource for production add-ins.</span></span>
* <span data-ttu-id="d49ab-118">https://appsforoffice.microsoft.com/lib/beta/hosted/office.js-Utilisez cette ressource lorsque vous essayez les fonctionnalités d’aperçu.</span><span class="sxs-lookup"><span data-stu-id="d49ab-118">https://appsforoffice.microsoft.com/lib/beta/hosted/office.js - use this resource when you're trying out preview features.</span></span>

<span data-ttu-id="d49ab-p103">Si vous utilisez [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), vous pouvez télécharger les [outils de développement Office](https://www.visualstudio.com/features/office-tools-vs.aspx) pour obtenir des modèles de projets qui incluent Office.js.  Vous pouvez également utiliser [nuget pour obtenir Office.js](https://www.nuget.org/packages/Microsoft.Office.js/).</span><span class="sxs-lookup"><span data-stu-id="d49ab-p103">If you're using [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), you can download the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) to get project templates that include Office.js.  You can also use [nuget to get Office.js](https://www.nuget.org/packages/Microsoft.Office.js/).</span></span>

<span data-ttu-id="d49ab-121">Si vous utilisez TypeScript et que vous avez npm, vous pouvez obtenir les définitions TypeScript en tapant `typings install office-js --ambient` dans votre interface de ligne de commande.</span><span class="sxs-lookup"><span data-stu-id="d49ab-121">If you use TypeScript and have npm, you can get the the TypeScript definitions by typing this in your command line interface: `typings install office-js --ambient`.</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="d49ab-122">Exécution de compléments Word</span><span class="sxs-lookup"><span data-stu-id="d49ab-122">Running Word add-ins</span></span>

<span data-ttu-id="d49ab-p104">Pour exécuter votre complément, utilisez un gestionnaire d’événements Office.initialize. Pour plus d’informations sur l’initialisation du complément, voir [Présentation de l’API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office) .</span><span class="sxs-lookup"><span data-stu-id="d49ab-p104">To run your add-in, use an Office.initialize event handler. For more information about add-in initialization, see [Understanding the API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office) .</span></span>

<span data-ttu-id="d49ab-125">Compléments qui ciblent 2016 Word ou exécuter en passant une fonction dans la méthode **Word.run()** .</span><span class="sxs-lookup"><span data-stu-id="d49ab-125">Add-ins that target Word 2016 or later execute by passing a function into the **Word.run()** method.</span></span> <span data-ttu-id="d49ab-126">La fonction passée dans la méthode **run** doit avoir un argument de contexte.</span><span class="sxs-lookup"><span data-stu-id="d49ab-126">The function passed into the **run** method must have a context argument.</span></span> <span data-ttu-id="d49ab-127">Cet [objet de contexte](/javascript/api/word/word.requestcontext) est différente de l’objet context que vous obtenez de l’objet Office, mais il est également utilisé pour interagir avec l’environnement d’exécution de Word.</span><span class="sxs-lookup"><span data-stu-id="d49ab-127">This [context object](/javascript/api/word/word.requestcontext) is different than the context object you get from the Office object, but it is also used to interact with the Word runtime environment.</span></span> <span data-ttu-id="d49ab-128">L’objet context donne accès au modèle objet Word JavaScript API.</span><span class="sxs-lookup"><span data-stu-id="d49ab-128">The context object provides access to the Word JavaScript API object model.</span></span> <span data-ttu-id="d49ab-129">L’exemple suivant montre comment initialiser et exécuter un mot complément à l’aide de la méthode **Word.run()** .</span><span class="sxs-lookup"><span data-stu-id="d49ab-129">The following example shows how to initialize and execute a Word add-in by using the **Word.run()** method.</span></span>

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a><span data-ttu-id="d49ab-130">Synchronisation de documents Word avec des objets de proxy de l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="d49ab-130">Synchronizing Word documents with Word JavaScript API proxy objects</span></span>

<span data-ttu-id="d49ab-p106">Le modèle objet de l’API JavaScript pour Word est associé de façon relativement libre aux objets dans Word. Les objets de l’API JavaScript pour Word sont des proxys pour des objets dans un document Word. Les actions effectuées sur les objets de proxy ne sont pas réalisées dans Word tant que l’état du document n’a pas été synchronisé. Inversement, l’état du document Word n’est pas répercuté sur les objets de proxy tant que l’état du document n’a pas été synchronisé. Pour synchroniser l’état du document, vous exécutez la méthode **context.sync()**. L’exemple suivant présente la création d’un objet Body de proxy et une file de commandes permettant de charger la propriété de texte sur l’objet Body de proxy, puis la synchronisation du corps dans le document Word avec l’objet de proxy correspondant à l’aide de la méthode **context.sync()**.</span><span class="sxs-lookup"><span data-stu-id="d49ab-p106">The Word JavaScript API object model is loosely coupled with the objects in Word. Word JavaScript API objects are proxies for objects in a Word document. Actions taken on proxy objects are not realized in Word until the document state has been synchronized. Conversely, the state of the Word document is not realized in the proxy objects until the document state has been synchronized. To synchronize the document state, you run the **context.sync()** method. The following example creates a proxy body object and a queued command to load the text property on the proxy body object, and uses the **context.sync()** method to synchronize the body of the Word document with the body proxy object.</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a><span data-ttu-id="d49ab-137">Exécution d’un lot de commandes</span><span class="sxs-lookup"><span data-stu-id="d49ab-137">Executing a batch of commands</span></span>

<span data-ttu-id="d49ab-p107">Les objets de proxy Word utilisent des méthodes pour accéder au modèle objet et le mettre à jour. Ces méthodes sont exécutées l’une après l’autre, dans l’ordre dans lequel elles ont incluses dans la file d’attente du lot. Toutes les commandes en attente dans le lot sont exécutées lorsque la méthode context.sync() est appelée.</span><span class="sxs-lookup"><span data-stu-id="d49ab-p107">The Word proxy objects have methods for accessing and updating the object model. These methods are executed sequentially in the order in which they were queued in the batch. All of the commands that are queued in the batch are executed when context.sync() is called.</span></span>

<span data-ttu-id="d49ab-p108">L’exemple suivant montre comment fonctionne la file d’attente de commandes. Lorsque la méthode **context.sync()** est appelée, la commande visant à charger le corps du texte est exécutée dans Word. C’est ensuite la commande visant à insérer du texte dans le corps de Word qui est appliquée. Les résultats sont alors renvoyés vers l’objet Body de proxy. La valeur de la propriété **body.text** dans l’API JavaScript pour Word est la valeur du corps du document de Word <u>avant</u> l’insertion du texte dans le document Word.</span><span class="sxs-lookup"><span data-stu-id="d49ab-p108">The following example shows how the command queue works. When **context.sync()** is called, the command to load the body text is executed in Word. Then, the command to insert text into the body in Word occurs. The results are then returned to the body proxy object. The value of the **body.text** property in the Word JavaScript API is the value of the Word document body <u>before</u> the text was inserted into Word document.</span></span>


```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    context.load(body, 'text');

    // Queue a command to insert text into the end of the Word document body.
    body.insertText('This is text inserted after loading the body.text property',
                    Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

## <a name="word-javascript-api-open-specifications"></a><span data-ttu-id="d49ab-146">Spécifications ouvertes Word JavaScript API</span><span class="sxs-lookup"><span data-stu-id="d49ab-146">Word JavaScript API open specifications</span></span>

<span data-ttu-id="d49ab-p109">Au fur et à mesure que nous concevons et développons de nouvelles API pour les compléments Word, nous les mettons à votre disposition sur notre page de [spécifications d’ouverture de l’API](../openspec.md) pour que vous puissiez fournir vos commentaires. Découvrez les nouvelles fonctionnalités dans le pipeline pour les API JavaScript pour Word et donnez votre avis sur nos spécifications de conception.</span><span class="sxs-lookup"><span data-stu-id="d49ab-p109">As we design and develop new APIs for Word add-ins, we'll make them available for your feedback on our [Open API specifications](../openspec.md) page. Find out what new features are in the pipeline for the Word JavaScript APIs, and provide your input on our design specifications.</span></span>

## <a name="word-javascript-api-reference"></a><span data-ttu-id="d49ab-149">Référence d’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="d49ab-149">Word JavaScript API reference</span></span>

<span data-ttu-id="d49ab-150">Pour plus d’informations sur l’API JavaScript de Word, consultez la [documentation de référence API JavaScript de Word](/javascript/api/word).</span><span class="sxs-lookup"><span data-stu-id="d49ab-150">For detailed information about the Word JavaScript API, see the [Word JavaScript API reference documentation](/javascript/api/word).</span></span>

## <a name="see-also"></a><span data-ttu-id="d49ab-151">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d49ab-151">See also</span></span>

* [<span data-ttu-id="d49ab-152">Présentation des compléments Word</span><span class="sxs-lookup"><span data-stu-id="d49ab-152">Word add-ins overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/word/word-add-ins-programming-overview)
* [<span data-ttu-id="d49ab-153">Vue d’ensemble de la plateforme des compléments Office</span><span class="sxs-lookup"><span data-stu-id="d49ab-153">Office Add-ins platform overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* [<span data-ttu-id="d49ab-154">Exemples de compléments Word sur GitHub</span><span class="sxs-lookup"><span data-stu-id="d49ab-154">Word add-in samples on GitHub</span></span>](https://github.com/OfficeDev?utf8=%E2%9C%93&q=Word)
