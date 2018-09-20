# <a name="functionfile-element"></a><span data-ttu-id="5544f-101">Élément FunctionFile</span><span class="sxs-lookup"><span data-stu-id="5544f-101">FunctionFile element</span></span>

<span data-ttu-id="5544f-p101">Spécifie le fichier de code source pour les opérations qu’un complément expose via les commandes de complément qui exécutent une fonction JavaScript au lieu d’afficher l’interface utilisateur. L’élément **FunctionFile** est un élément enfant de [DesktopFormFactor](desktopformfactor.md) ou de [MobileFormFactor](mobileformfactor.md). L’attribut **resid** de l’élément **FunctionFile** est défini sur la valeur de l’attribut **id** d’un élément **Url** dans l’élément **Resources** contenant l’URL d’un fichier HTML qui contient ou charge toutes les fonctions JavaScript utilisées par les boutons de commande de complément sans interface utilisateur, telles que définies par l’élément [Control](control.md).</span><span class="sxs-lookup"><span data-stu-id="5544f-p101">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI. The  **FunctionFile** element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md). The **resid** attribute of the **FunctionFile** element is set to the value of the **id** attribute of a **Url** element in the **Resources** element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="5544f-105">Voici un exemple de l’élément **FunctionFile** .</span><span class="sxs-lookup"><span data-stu-id="5544f-105">The following is an example of the  **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="5544f-106">Le code JavaScript dans le fichier HTML indiqué par l’élément **FunctionFile** doit appeler `Office.initialize` et définir des fonctions nommées qui acceptent un seul paramètre : `event`.</span><span class="sxs-lookup"><span data-stu-id="5544f-106">The JavaScript in the HTML file indicated by the  **FunctionFile** element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="5544f-107">Les fonctions doivent utiliser le `item.notificationMessages` API pour indiquer la progression, réussite ou Échec de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="5544f-107">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="5544f-108">Elle doit également appeler `event.completed` après avoir terminé l’exécution.</span><span class="sxs-lookup"><span data-stu-id="5544f-108">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="5544f-109">Le nom des fonctions sont utilisés dans l’élément **suivantes** pour les boutons sans interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="5544f-109">The name of the functions are used in the **FunctionName** element for UI-less buttons.</span></span>

<span data-ttu-id="5544f-110">Vous trouverez ci-dessous un exemple d’un fichier HTML définissant une fonction **trackMessage**.</span><span class="sxs-lookup"><span data-stu-id="5544f-110">The following is an example of an HTML file defining a **trackMessage** function.</span></span>

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

<span data-ttu-id="5544f-111">Le code suivant montre comment implémenter la fonction utilisée par **suivantes**.</span><span class="sxs-lookup"><span data-stu-id="5544f-111">The following code shows how to implement the function used by  **FunctionName**.</span></span>

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> <span data-ttu-id="5544f-112">L’appel vers **event.completed** signale que vous avez correctement géré l’événement.</span><span class="sxs-lookup"><span data-stu-id="5544f-112">The call to  **event.completed** signals that you have successfully handled the event.</span></span> <span data-ttu-id="5544f-113">Lorsqu’une fonction est appelée à plusieurs reprises, par exemple, lorsque l’utilisateur clique plusieurs fois sur une même commande de complément, tous les événements sont automatiquement mis en file d’attente.</span><span class="sxs-lookup"><span data-stu-id="5544f-113">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="5544f-114">Le premier événement s’exécute automatiquement, tandis que les autres événements restent dans la file d’attente.</span><span class="sxs-lookup"><span data-stu-id="5544f-114">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="5544f-115">Lorsque votre fonction appelle **event.completed**, l’appel suivant de cette fonction s’exécute.</span><span class="sxs-lookup"><span data-stu-id="5544f-115">When your function calls **event.completed**, the next queued call to that function runs.</span></span> <span data-ttu-id="5544f-116">Vous devez appeler **event.completed**; Sinon, votre fonction ne fonctionnera pas.</span><span class="sxs-lookup"><span data-stu-id="5544f-116">You must call **event.completed**; otherwise your function will not run.</span></span>