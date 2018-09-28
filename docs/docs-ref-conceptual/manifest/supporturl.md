# <a name="supporturl-element"></a><span data-ttu-id="05ab1-101">SupportUrl, élément</span><span class="sxs-lookup"><span data-stu-id="05ab1-101">SupportUrl element</span></span>

<span data-ttu-id="05ab1-102">Spécifie l’URL d’une page qui fournit des informations de prise en charge pour votre complément.</span><span class="sxs-lookup"><span data-stu-id="05ab1-102">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="05ab1-103">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="05ab1-103">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="05ab1-104">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="05ab1-104">Contained in</span></span>

[<span data-ttu-id="05ab1-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="05ab1-105">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="05ab1-106">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="05ab1-106">Can contain</span></span>

|  <span data-ttu-id="05ab1-107">Élément</span><span class="sxs-lookup"><span data-stu-id="05ab1-107">Element</span></span> | <span data-ttu-id="05ab1-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="05ab1-108">Required</span></span> | <span data-ttu-id="05ab1-109">Description</span><span class="sxs-lookup"><span data-stu-id="05ab1-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="05ab1-110">Override</span><span class="sxs-lookup"><span data-stu-id="05ab1-110">Override</span></span>](override.md)   | <span data-ttu-id="05ab1-111">Non</span><span class="sxs-lookup"><span data-stu-id="05ab1-111">No</span></span> | <span data-ttu-id="05ab1-112">Spécifie le paramètre pour les URL de paramètres régionaux supplémentaires</span><span class="sxs-lookup"><span data-stu-id="05ab1-112">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="05ab1-113">Attributs</span><span class="sxs-lookup"><span data-stu-id="05ab1-113">Attributes</span></span>

|<span data-ttu-id="05ab1-114">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="05ab1-114">**Attribute**</span></span>|<span data-ttu-id="05ab1-115">**Type**</span><span class="sxs-lookup"><span data-stu-id="05ab1-115">**Type**</span></span>|<span data-ttu-id="05ab1-116">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="05ab1-116">**Required**</span></span>|<span data-ttu-id="05ab1-117">**Description**</span><span class="sxs-lookup"><span data-stu-id="05ab1-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="05ab1-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="05ab1-118">DefaultValue</span></span>|<span data-ttu-id="05ab1-119">URL</span><span class="sxs-lookup"><span data-stu-id="05ab1-119">URL</span></span>|<span data-ttu-id="05ab1-120">obligatoire</span><span class="sxs-lookup"><span data-stu-id="05ab1-120">required</span></span>|<span data-ttu-id="05ab1-121">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="05ab1-121">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
