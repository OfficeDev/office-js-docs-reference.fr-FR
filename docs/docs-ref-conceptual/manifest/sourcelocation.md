# <a name="sourcelocation-element"></a><span data-ttu-id="934e3-101">Élément SourceLocation</span><span class="sxs-lookup"><span data-stu-id="934e3-101">SourceLocation element</span></span>

<span data-ttu-id="934e3-p101">Spécifie les emplacements des fichiers source pour votre complément Office sous forme d’URL comprenant entre 1 et 2 018 caractères. L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.</span><span class="sxs-lookup"><span data-stu-id="934e3-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="934e3-104">**Type de complément :** Application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="934e3-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="934e3-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="934e3-105">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="934e3-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="934e3-106">Contained in</span></span>

- <span data-ttu-id="934e3-107">[DefaultSettings](defaultsettings.md) (compléments de contenu et de volet Office)</span><span class="sxs-lookup"><span data-stu-id="934e3-107">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="934e3-108">[FormSettings](formsettings.md) (compléments de messagerie)</span><span class="sxs-lookup"><span data-stu-id="934e3-108">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="934e3-109">[ExtensionPoint](extensionpoint.md) (compléments de messagerie contextuels)</span><span class="sxs-lookup"><span data-stu-id="934e3-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="934e3-110">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="934e3-110">Can contain</span></span>

[<span data-ttu-id="934e3-111">Override</span><span class="sxs-lookup"><span data-stu-id="934e3-111">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="934e3-112">Attributs</span><span class="sxs-lookup"><span data-stu-id="934e3-112">Attributes</span></span>

|<span data-ttu-id="934e3-113">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="934e3-113">**Attribute**</span></span>|<span data-ttu-id="934e3-114">**Type**</span><span class="sxs-lookup"><span data-stu-id="934e3-114">**Type**</span></span>|<span data-ttu-id="934e3-115">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="934e3-115">**Required**</span></span>|<span data-ttu-id="934e3-116">**Description**</span><span class="sxs-lookup"><span data-stu-id="934e3-116">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="934e3-117">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="934e3-117">DefaultValue</span></span>|<span data-ttu-id="934e3-118">URL</span><span class="sxs-lookup"><span data-stu-id="934e3-118">URL</span></span>|<span data-ttu-id="934e3-119">obligatoire</span><span class="sxs-lookup"><span data-stu-id="934e3-119">required</span></span>|<span data-ttu-id="934e3-120">Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="934e3-120">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
