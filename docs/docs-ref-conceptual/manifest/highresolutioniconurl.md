# <a name="highresolutioniconurl-element"></a><span data-ttu-id="6cdf2-101">HighResolutionIconUrl, élément</span><span class="sxs-lookup"><span data-stu-id="6cdf2-101">HighResolutionIconUrl element</span></span>

<span data-ttu-id="6cdf2-102">Spécifie l’URL de l’image qui est utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store sur les écrans à haute résolution (DPI).</span><span class="sxs-lookup"><span data-stu-id="6cdf2-102">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="6cdf2-103">**Type de complément :** Application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="6cdf2-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6cdf2-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="6cdf2-104">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="6cdf2-105">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="6cdf2-105">Can contain</span></span>

[<span data-ttu-id="6cdf2-106">Override</span><span class="sxs-lookup"><span data-stu-id="6cdf2-106">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="6cdf2-107">Attributs</span><span class="sxs-lookup"><span data-stu-id="6cdf2-107">Attributes</span></span>

|<span data-ttu-id="6cdf2-108">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="6cdf2-108">**Attribute**</span></span>|<span data-ttu-id="6cdf2-109">**Type**</span><span class="sxs-lookup"><span data-stu-id="6cdf2-109">**Type**</span></span>|<span data-ttu-id="6cdf2-110">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="6cdf2-110">**Required**</span></span>|<span data-ttu-id="6cdf2-111">**Description**</span><span class="sxs-lookup"><span data-stu-id="6cdf2-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="6cdf2-112">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="6cdf2-112">DefaultValue</span></span>|<span data-ttu-id="6cdf2-113">chaîne (URL)</span><span class="sxs-lookup"><span data-stu-id="6cdf2-113">string (URL)</span></span>|<span data-ttu-id="6cdf2-114">obligatoire</span><span class="sxs-lookup"><span data-stu-id="6cdf2-114">required</span></span>|<span data-ttu-id="6cdf2-115">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="6cdf2-115">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="6cdf2-116">Remarques</span><span class="sxs-lookup"><span data-stu-id="6cdf2-116">Remarks</span></span>

<span data-ttu-id="6cdf2-p101">Pour un complément de messagerie, l’icône apparaît dans l’interface utilisateur, sous **Fichier**  >  **Gérer les compléments**. Pour un complément de contenu ou de volet Office, l’icône apparaît dans l’interface utilisateur, sous **Insérer**  >  **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="6cdf2-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="6cdf2-119">L’image doit être dans l’un des formats de fichier suivants, avec une résolution recommandée de 64 x 64 pixels : GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="6cdf2-119">The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF.</span></span> <span data-ttu-id="6cdf2-120">Pour plus d’informations, voir la section _créer une identité visuelle cohérente pour votre application_ de [créer des annonces efficaces dans AppSource et dans Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="6cdf2-120">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span></span>
