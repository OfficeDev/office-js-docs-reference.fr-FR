# <a name="override-element"></a><span data-ttu-id="be76a-101">Élément Override</span><span class="sxs-lookup"><span data-stu-id="be76a-101">Override element</span></span>

<span data-ttu-id="be76a-102">Fournit une manière de spécifier la valeur d’un paramètre pour d’autres paramètres régionaux.</span><span class="sxs-lookup"><span data-stu-id="be76a-102">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="be76a-103">**Type de complément :** Application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="be76a-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="be76a-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="be76a-104">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="be76a-105">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="be76a-105">Contained in</span></span>

|<span data-ttu-id="be76a-106">**Élément**</span><span class="sxs-lookup"><span data-stu-id="be76a-106">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="be76a-107">CitationText</span><span class="sxs-lookup"><span data-stu-id="be76a-107">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="be76a-108">Description</span><span class="sxs-lookup"><span data-stu-id="be76a-108">Description</span></span>](description.md)|
|[<span data-ttu-id="be76a-109">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="be76a-109">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="be76a-110">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="be76a-110">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="be76a-111">DisplayName</span><span class="sxs-lookup"><span data-stu-id="be76a-111">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="be76a-112">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="be76a-112">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="be76a-113">IconUrl</span><span class="sxs-lookup"><span data-stu-id="be76a-113">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="be76a-114">QueryUri</span><span class="sxs-lookup"><span data-stu-id="be76a-114">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="be76a-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="be76a-115">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="be76a-116">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="be76a-116">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="be76a-117">Attributs</span><span class="sxs-lookup"><span data-stu-id="be76a-117">Attributes</span></span>

|<span data-ttu-id="be76a-118">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="be76a-118">**Attribute**</span></span>|<span data-ttu-id="be76a-119">**Type**</span><span class="sxs-lookup"><span data-stu-id="be76a-119">**Type**</span></span>|<span data-ttu-id="be76a-120">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="be76a-120">**Required**</span></span>|<span data-ttu-id="be76a-121">**Description**</span><span class="sxs-lookup"><span data-stu-id="be76a-121">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="be76a-122">Locale</span><span class="sxs-lookup"><span data-stu-id="be76a-122">Locale</span></span>|<span data-ttu-id="be76a-123">string</span><span class="sxs-lookup"><span data-stu-id="be76a-123">string</span></span>|<span data-ttu-id="be76a-124">obligatoire</span><span class="sxs-lookup"><span data-stu-id="be76a-124">required</span></span>|<span data-ttu-id="be76a-125">Spécifie le nom de culture des paramètres régionaux pour ce remplacement au format de balise de langue BCP 47, comme `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="be76a-125">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="be76a-126">Valeur</span><span class="sxs-lookup"><span data-stu-id="be76a-126">Value</span></span>|<span data-ttu-id="be76a-127">string</span><span class="sxs-lookup"><span data-stu-id="be76a-127">string</span></span>|<span data-ttu-id="be76a-128">obligatoire</span><span class="sxs-lookup"><span data-stu-id="be76a-128">required</span></span>|<span data-ttu-id="be76a-129">Spécifie la valeur du paramètre exprimée pour les paramètres régionaux spécifiés.</span><span class="sxs-lookup"><span data-stu-id="be76a-129">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="be76a-130">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="be76a-130">See also</span></span>

- [<span data-ttu-id="be76a-131">Localisation des compléments Office</span><span class="sxs-lookup"><span data-stu-id="be76a-131">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
