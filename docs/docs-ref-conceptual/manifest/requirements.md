# <a name="requirements-element"></a><span data-ttu-id="133e1-101">Élément Requirements</span><span class="sxs-lookup"><span data-stu-id="133e1-101">Requirements element</span></span>

<span data-ttu-id="133e1-102">Spécifie l’ensemble minimal des conditions requises de l’API JavaScript pour Office ([ensembles des conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) et/ou méthodes) que votre complément Office doit activer.</span><span class="sxs-lookup"><span data-stu-id="133e1-102">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="133e1-103">**Type de complément :** Application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="133e1-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="133e1-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="133e1-104">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="133e1-105">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="133e1-105">Contained in</span></span>

[<span data-ttu-id="133e1-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="133e1-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="133e1-107">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="133e1-107">Can contain</span></span>

|<span data-ttu-id="133e1-108">**Élément**</span><span class="sxs-lookup"><span data-stu-id="133e1-108">**Element**</span></span>|<span data-ttu-id="133e1-109">**Content**</span><span class="sxs-lookup"><span data-stu-id="133e1-109">**Content**</span></span>|<span data-ttu-id="133e1-110">**Courrier**</span><span class="sxs-lookup"><span data-stu-id="133e1-110">**Mail**</span></span>|<span data-ttu-id="133e1-111">**Volet Office**</span><span class="sxs-lookup"><span data-stu-id="133e1-111">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="133e1-112">Ensembles</span><span class="sxs-lookup"><span data-stu-id="133e1-112">Sets</span></span>](sets.md)|<span data-ttu-id="133e1-113">x</span><span class="sxs-lookup"><span data-stu-id="133e1-113">x</span></span>|<span data-ttu-id="133e1-114">x</span><span class="sxs-lookup"><span data-stu-id="133e1-114">x</span></span>|<span data-ttu-id="133e1-115">x</span><span class="sxs-lookup"><span data-stu-id="133e1-115">x</span></span>|
|[<span data-ttu-id="133e1-116">Méthodes</span><span class="sxs-lookup"><span data-stu-id="133e1-116">Methods</span></span>](methods.md)|<span data-ttu-id="133e1-117">x</span><span class="sxs-lookup"><span data-stu-id="133e1-117">x</span></span>||<span data-ttu-id="133e1-118">x</span><span class="sxs-lookup"><span data-stu-id="133e1-118">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="133e1-119">Remarques</span><span class="sxs-lookup"><span data-stu-id="133e1-119">Remarks</span></span>

<span data-ttu-id="133e1-120">Pour plus d’informations sur les ensembles de ressources, voir [définit les versions d’Office et de la spécification](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="133e1-120">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

