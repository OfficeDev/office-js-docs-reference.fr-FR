# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="cc549-101">Ensembles de conditions requises de l’API JavaScript pour OneNote</span><span class="sxs-lookup"><span data-stu-id="cc549-101">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="cc549-102">Les ensembles de conditions requises sont des groupes nommés de membres d’API.</span><span class="sxs-lookup"><span data-stu-id="cc549-102">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="cc549-103">Compléments Office utilisent ensembles spécifiés dans le manifeste ou une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API nécessitant un complément.</span><span class="sxs-lookup"><span data-stu-id="cc549-103">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="cc549-104">Pour plus d’informations, voir [définit les versions d’Office et de la spécification](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="cc549-104">For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="cc549-105">Le tableau suivant répertorie les ensembles de conditions requises pour OneNote, les applications hôtes Office qui prennent en charge ces conditions et les numéros de version ou la date de disponibilité.</span><span class="sxs-lookup"><span data-stu-id="cc549-105">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="cc549-106">Ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="cc549-106">Requirement set</span></span>  |  <span data-ttu-id="cc549-107">Office Online</span><span class="sxs-lookup"><span data-stu-id="cc549-107">Office Online</span></span> | 
|:-----|:-----|
| <span data-ttu-id="cc549-108">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="cc549-108">OneNoteApi 1.1</span></span>  | <span data-ttu-id="cc549-109">Septembre 2016</span><span class="sxs-lookup"><span data-stu-id="cc549-109">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="cc549-110">Ensembles de conditions requises des API communes pour Office</span><span class="sxs-lookup"><span data-stu-id="cc549-110">Office common API requirement sets</span></span>

<span data-ttu-id="cc549-111">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="cc549-111">For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="cc549-112">API JavaScript pour OneNote 1.1</span><span class="sxs-lookup"><span data-stu-id="cc549-112">OneNote JavaScript API 1.1</span></span> 

<span data-ttu-id="cc549-113">L’API JavaScript 1.1 pour OneNote est la première version de l’API.</span><span class="sxs-lookup"><span data-stu-id="cc549-113">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="cc549-114">Pour plus d’informations sur l’API, consultez la [vue d’ensemble de la programmation OneNote JavaScript API](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span><span class="sxs-lookup"><span data-stu-id="cc549-114">For details about the API, see the [OneNote JavaScript API programming overview](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="cc549-115">Vérification de la prise en charge d’un ensemble de conditions requises à l’exécution</span><span class="sxs-lookup"><span data-stu-id="cc549-115">Runtime requirement support check</span></span>

<span data-ttu-id="cc549-116">Lors de l’exécution, les compléments peuvent vérifier si un hôte particulier prend en charge un ensemble de conditions requises d’API en procédant comme suit :</span><span class="sxs-lookup"><span data-stu-id="cc549-116">During the runtime, add-ins can check if a particular host supports an API requirement set by doing the following-check:</span></span> 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="cc549-117">Vérification de la prise en charge des conditions requises basée sur le manifeste</span><span class="sxs-lookup"><span data-stu-id="cc549-117">Manifest-based requirement support check</span></span>

<span data-ttu-id="cc549-p103">Utilisez l’élément Conditions requises dans le manifeste du complément pour spécifier des ensembles de conditions requises essentiels ou des membres d’API que votre complément doit utiliser. Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l’élément Conditions requises, le complément ne s’exécute pas dans cet hôte ou cette plateforme et ne s’affiche pas dans Mes compléments.</span><span class="sxs-lookup"><span data-stu-id="cc549-p103">Use the Requirements element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the Requirements element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="cc549-120">Cet exemple de code illustre un complément qui se charge dans toutes les applications hôtes Office qui prennent en charge l’ensemble de conditions requises OneNoteApi, version 1.1.</span><span class="sxs-lookup"><span data-stu-id="cc549-120">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a><span data-ttu-id="cc549-121">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="cc549-121">See also</span></span>

- [<span data-ttu-id="cc549-122">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="cc549-122">Office versions and requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="cc549-123">Spécification des exigences en matière d’hôtes Office et d’API</span><span class="sxs-lookup"><span data-stu-id="cc549-123">Specify Office hosts and API requirements</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="cc549-124">Manifeste XML des compléments Office</span><span class="sxs-lookup"><span data-stu-id="cc549-124">Office Add-ins XML manifest</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
