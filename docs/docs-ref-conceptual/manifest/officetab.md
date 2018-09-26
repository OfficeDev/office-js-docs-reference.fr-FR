# <a name="officetab-element"></a><span data-ttu-id="1fdcc-101">Élément OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1fdcc-101">OfficeTab element</span></span>

<span data-ttu-id="1fdcc-p101">Définit l’onglet du ruban sur lequel votre commande de complément s’affiche. Il peut s’agir de l’onglet par défaut (soit **Accueil**, **Message** ou **Réunion**), ou d’un onglet personnalisé défini par le complément. Cet élément est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="1fdcc-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1fdcc-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="1fdcc-105">Child elements</span></span>

|  <span data-ttu-id="1fdcc-106">Élément</span><span class="sxs-lookup"><span data-stu-id="1fdcc-106">Element</span></span> |  <span data-ttu-id="1fdcc-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="1fdcc-107">Required</span></span>  |  <span data-ttu-id="1fdcc-108">Description</span><span class="sxs-lookup"><span data-stu-id="1fdcc-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1fdcc-109">Group</span><span class="sxs-lookup"><span data-stu-id="1fdcc-109">Group</span></span>      | <span data-ttu-id="1fdcc-110">Oui</span><span class="sxs-lookup"><span data-stu-id="1fdcc-110">Yes</span></span> |  <span data-ttu-id="1fdcc-p102">Définit un groupe de commandes. Vous ne pouvez ajouter qu’un seul groupe par complément à l’onglet par défaut.</span><span class="sxs-lookup"><span data-stu-id="1fdcc-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="1fdcc-113">Les valeurs suivantes sont des valeurs `id` d’onglet valides par hôte.</span><span class="sxs-lookup"><span data-stu-id="1fdcc-113">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="1fdcc-114">Valeurs en **gras** sont pris en charge dans les postes de travail et en ligne (par exemple, Word 2016 ou version ultérieure pour Windows et Word en ligne).</span><span class="sxs-lookup"><span data-stu-id="1fdcc-114">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="1fdcc-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="1fdcc-115">Outlook</span></span>

- <span data-ttu-id="1fdcc-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="1fdcc-117">Word</span><span class="sxs-lookup"><span data-stu-id="1fdcc-117">Word</span></span>

- <span data-ttu-id="1fdcc-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-118">**TabHome**</span></span>
- <span data-ttu-id="1fdcc-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-119">**TabInsert**</span></span>
- <span data-ttu-id="1fdcc-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="1fdcc-120">TabWordDesign</span></span>
- <span data-ttu-id="1fdcc-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="1fdcc-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="1fdcc-122">TabReferences</span></span>
- <span data-ttu-id="1fdcc-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="1fdcc-123">TabMailings</span></span>
- <span data-ttu-id="1fdcc-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="1fdcc-124">TabReviewWord</span></span>
- <span data-ttu-id="1fdcc-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-125">**TabView**</span></span>
- <span data-ttu-id="1fdcc-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="1fdcc-126">TabDeveloper</span></span>
- <span data-ttu-id="1fdcc-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="1fdcc-127">TabAddIns</span></span>
- <span data-ttu-id="1fdcc-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="1fdcc-128">TabBlogPost</span></span>
- <span data-ttu-id="1fdcc-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="1fdcc-129">TabBlogInsert</span></span>
- <span data-ttu-id="1fdcc-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="1fdcc-130">TabPrintPreview</span></span>
- <span data-ttu-id="1fdcc-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="1fdcc-131">TabOutlining</span></span>
- <span data-ttu-id="1fdcc-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="1fdcc-132">TabConflicts</span></span>
- <span data-ttu-id="1fdcc-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="1fdcc-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="1fdcc-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="1fdcc-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="1fdcc-135">Excel</span><span class="sxs-lookup"><span data-stu-id="1fdcc-135">Excel</span></span>

- <span data-ttu-id="1fdcc-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-136">**TabHome**</span></span>
- <span data-ttu-id="1fdcc-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-137">**TabInsert**</span></span>
- <span data-ttu-id="1fdcc-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="1fdcc-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="1fdcc-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="1fdcc-139">TabFormulas</span></span>
- <span data-ttu-id="1fdcc-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-140">**TabData**</span></span>
- <span data-ttu-id="1fdcc-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-141">**TabReview**</span></span>
- <span data-ttu-id="1fdcc-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-142">**TabView**</span></span>
- <span data-ttu-id="1fdcc-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="1fdcc-143">TabDeveloper</span></span>
- <span data-ttu-id="1fdcc-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="1fdcc-144">TabAddIns</span></span>
- <span data-ttu-id="1fdcc-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="1fdcc-145">TabPrintPreview</span></span>
- <span data-ttu-id="1fdcc-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="1fdcc-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="1fdcc-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1fdcc-147">PowerPoint</span></span>

- <span data-ttu-id="1fdcc-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-148">**TabHome**</span></span>
- <span data-ttu-id="1fdcc-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-149">**TabInsert**</span></span>
- <span data-ttu-id="1fdcc-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-150">**TabDesign**</span></span>
- <span data-ttu-id="1fdcc-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-151">**TabTransitions**</span></span>
- <span data-ttu-id="1fdcc-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-152">**TabAnimations**</span></span>
- <span data-ttu-id="1fdcc-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="1fdcc-153">TabSlideShow</span></span>
- <span data-ttu-id="1fdcc-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="1fdcc-154">TabReview</span></span>
- <span data-ttu-id="1fdcc-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-155">**TabView**</span></span>
- <span data-ttu-id="1fdcc-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="1fdcc-156">TabDeveloper</span></span>
- <span data-ttu-id="1fdcc-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="1fdcc-157">TabAddIns</span></span>
- <span data-ttu-id="1fdcc-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="1fdcc-158">TabPrintPreview</span></span>
- <span data-ttu-id="1fdcc-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="1fdcc-159">TabMerge</span></span>
- <span data-ttu-id="1fdcc-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="1fdcc-160">TabGrayscale</span></span>
- <span data-ttu-id="1fdcc-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="1fdcc-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="1fdcc-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="1fdcc-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="1fdcc-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="1fdcc-163">TabSlideMaster</span></span>
- <span data-ttu-id="1fdcc-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="1fdcc-164">TabHandoutMaster</span></span>
- <span data-ttu-id="1fdcc-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="1fdcc-165">TabNotesMaster</span></span>
- <span data-ttu-id="1fdcc-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="1fdcc-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="1fdcc-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="1fdcc-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="1fdcc-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="1fdcc-168">OneNote</span></span>

- <span data-ttu-id="1fdcc-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-169">**TabHome**</span></span>
- <span data-ttu-id="1fdcc-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-170">**TabInsert**</span></span>
- <span data-ttu-id="1fdcc-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="1fdcc-171">**TabView**</span></span>
- <span data-ttu-id="1fdcc-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="1fdcc-172">TabDeveloper</span></span>
- <span data-ttu-id="1fdcc-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="1fdcc-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="1fdcc-174">Group</span><span class="sxs-lookup"><span data-stu-id="1fdcc-174">Group</span></span>

<span data-ttu-id="1fdcc-p104">Groupe de points d’extension d’interface utilisateur dans un onglet. Un groupe peut contenir jusqu’à six contrôles. L’attribut **id** est requis et chaque **id** doit être unique au sein du manifeste. L’**ID** est une chaîne avec un maximum de 125 caractères. Voir l’[élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="1fdcc-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="1fdcc-179">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1fdcc-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
