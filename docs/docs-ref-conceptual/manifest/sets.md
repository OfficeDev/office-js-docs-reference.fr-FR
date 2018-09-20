# <a name="sets-element"></a><span data-ttu-id="c9e5b-101">Élément Sets</span><span class="sxs-lookup"><span data-stu-id="c9e5b-101">Sets element</span></span>

<span data-ttu-id="c9e5b-102">Spécifie le sous-ensemble minimal de l’API JavaScript pour Office nécessaire à l’activation de votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="c9e5b-102">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="c9e5b-103">**Type de complément :** Application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="c9e5b-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c9e5b-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="c9e5b-104">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="c9e5b-105">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="c9e5b-105">Contained in</span></span>

[<span data-ttu-id="c9e5b-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c9e5b-106">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="c9e5b-107">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="c9e5b-107">Can contain</span></span>

[<span data-ttu-id="c9e5b-108">Ensemble</span><span class="sxs-lookup"><span data-stu-id="c9e5b-108">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="c9e5b-109">Attributs</span><span class="sxs-lookup"><span data-stu-id="c9e5b-109">Attributes</span></span>

|<span data-ttu-id="c9e5b-110">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="c9e5b-110">**Attribute**</span></span>|<span data-ttu-id="c9e5b-111">**Type**</span><span class="sxs-lookup"><span data-stu-id="c9e5b-111">**Type**</span></span>|<span data-ttu-id="c9e5b-112">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="c9e5b-112">**Required**</span></span>|<span data-ttu-id="c9e5b-113">**Description**</span><span class="sxs-lookup"><span data-stu-id="c9e5b-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="c9e5b-114">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="c9e5b-114">DefaultMinVersion</span></span>|<span data-ttu-id="c9e5b-115">chaîne</span><span class="sxs-lookup"><span data-stu-id="c9e5b-115">string</span></span>|<span data-ttu-id="c9e5b-116">facultatif</span><span class="sxs-lookup"><span data-stu-id="c9e5b-116">optional</span></span>|<span data-ttu-id="c9e5b-p101">Spécifie la valeur de l’attribut **MinVersion** par défaut pour tous les éléments [Set](set.md) enfants. La valeur par défaut est « 1.1 ».</span><span class="sxs-lookup"><span data-stu-id="c9e5b-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="c9e5b-119">Remarques</span><span class="sxs-lookup"><span data-stu-id="c9e5b-119">Remarks</span></span>

<span data-ttu-id="c9e5b-120">Pour plus d’informations sur les ensembles de ressources, voir [définit les versions d’Office et de la spécification](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="c9e5b-120">For more information about requirement sets, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="c9e5b-121">Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **Sets**, voir l’article relatif à la [définition de l’élément Requirements dans le manifeste](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="c9e5b-121">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

