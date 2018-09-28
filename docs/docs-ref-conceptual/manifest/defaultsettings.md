# <a name="defaultsettings-element"></a><span data-ttu-id="1da74-101">Élément DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="1da74-101">DefaultSettings element</span></span>

<span data-ttu-id="1da74-102">Spécifie l’emplacement de la source par défaut et d’autres paramètres par défaut pour le contenu ou le complément de volet de tâches.</span><span class="sxs-lookup"><span data-stu-id="1da74-102">Specifies the default source location and other default settings for your content or task pane add-in.</span></span>

<span data-ttu-id="1da74-103">**Type de complément :** Application de contenu et de volet Office</span><span class="sxs-lookup"><span data-stu-id="1da74-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="1da74-104">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="1da74-104">Syntax</span></span>

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a><span data-ttu-id="1da74-105">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="1da74-105">Contained in</span></span>

[<span data-ttu-id="1da74-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="1da74-106">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="1da74-107">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="1da74-107">Can contain</span></span>

|<span data-ttu-id="1da74-108">**Élément**</span><span class="sxs-lookup"><span data-stu-id="1da74-108">**Element**</span></span>|<span data-ttu-id="1da74-109">**Content**</span><span class="sxs-lookup"><span data-stu-id="1da74-109">**Content**</span></span>|<span data-ttu-id="1da74-110">**Courrier**</span><span class="sxs-lookup"><span data-stu-id="1da74-110">**Mail**</span></span>|<span data-ttu-id="1da74-111">**Volet Office**</span><span class="sxs-lookup"><span data-stu-id="1da74-111">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="1da74-112">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="1da74-112">SourceLocation</span></span>](sourcelocation.md)|<span data-ttu-id="1da74-113">x</span><span class="sxs-lookup"><span data-stu-id="1da74-113">x</span></span>||<span data-ttu-id="1da74-114">x</span><span class="sxs-lookup"><span data-stu-id="1da74-114">x</span></span>|
|[<span data-ttu-id="1da74-115">RequestedWidth</span><span class="sxs-lookup"><span data-stu-id="1da74-115">RequestedWidth</span></span>](requestedwidth.md)|<span data-ttu-id="1da74-116">x</span><span class="sxs-lookup"><span data-stu-id="1da74-116">x</span></span>|||
|[<span data-ttu-id="1da74-117">RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="1da74-117">RequestedHeight</span></span>](requestedheight.md)|<span data-ttu-id="1da74-118">x</span><span class="sxs-lookup"><span data-stu-id="1da74-118">x</span></span>|||

## <a name="remarks"></a><span data-ttu-id="1da74-119">Remarques</span><span class="sxs-lookup"><span data-stu-id="1da74-119">Remarks</span></span>

<span data-ttu-id="1da74-120">L’emplacement source et les autres paramètres de l’élément **DefaultSettings** s’appliquent uniquement aux compléments de volet Office et de contenu. Pour les compléments de messagerie, vous spécifiez les emplacements par défaut pour les fichiers sources et d’autres paramètres par défaut dans l’élément [FormSettings](formsettings.md).</span><span class="sxs-lookup"><span data-stu-id="1da74-120">The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](formsettings.md) element.</span></span>

