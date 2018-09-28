# <a name="officetab-element"></a><span data-ttu-id="34207-101">Élément OfficeTab</span><span class="sxs-lookup"><span data-stu-id="34207-101">OfficeTab element</span></span>

<span data-ttu-id="34207-p101">Définit l’onglet du ruban sur lequel votre commande de complément s’affiche. Il peut s’agir de l’onglet par défaut (soit **Accueil**, **Message** ou **Réunion**), ou d’un onglet personnalisé défini par le complément. Cet élément est obligatoire.</span><span class="sxs-lookup"><span data-stu-id="34207-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="34207-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="34207-105">Child elements</span></span>

|  <span data-ttu-id="34207-106">Élément</span><span class="sxs-lookup"><span data-stu-id="34207-106">Element</span></span> |  <span data-ttu-id="34207-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="34207-107">Required</span></span>  |  <span data-ttu-id="34207-108">Description</span><span class="sxs-lookup"><span data-stu-id="34207-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="34207-109">Group</span><span class="sxs-lookup"><span data-stu-id="34207-109">Group</span></span>      | <span data-ttu-id="34207-110">Oui</span><span class="sxs-lookup"><span data-stu-id="34207-110">Yes</span></span> |  <span data-ttu-id="34207-p102">Définit un groupe de commandes. Vous ne pouvez ajouter qu’un seul groupe par complément à l’onglet par défaut.</span><span class="sxs-lookup"><span data-stu-id="34207-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="34207-113">Les valeurs suivantes sont des valeurs `id` d’onglet valides par hôte.</span><span class="sxs-lookup"><span data-stu-id="34207-113">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="34207-114">Valeurs en **gras** sont pris en charge dans les postes de travail et en ligne (par exemple, Word 2016 ou version ultérieure pour Windows et Word en ligne).</span><span class="sxs-lookup"><span data-stu-id="34207-114">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="34207-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="34207-115">Outlook</span></span>

- <span data-ttu-id="34207-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="34207-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="34207-117">Word</span><span class="sxs-lookup"><span data-stu-id="34207-117">Word</span></span>

- <span data-ttu-id="34207-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="34207-118">**TabHome**</span></span>
- <span data-ttu-id="34207-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="34207-119">**TabInsert**</span></span>
- <span data-ttu-id="34207-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="34207-120">TabWordDesign</span></span>
- <span data-ttu-id="34207-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="34207-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="34207-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="34207-122">TabReferences</span></span>
- <span data-ttu-id="34207-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="34207-123">TabMailings</span></span>
- <span data-ttu-id="34207-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="34207-124">TabReviewWord</span></span>
- <span data-ttu-id="34207-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="34207-125">**TabView**</span></span>
- <span data-ttu-id="34207-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="34207-126">TabDeveloper</span></span>
- <span data-ttu-id="34207-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="34207-127">TabAddIns</span></span>
- <span data-ttu-id="34207-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="34207-128">TabBlogPost</span></span>
- <span data-ttu-id="34207-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="34207-129">TabBlogInsert</span></span>
- <span data-ttu-id="34207-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="34207-130">TabPrintPreview</span></span>
- <span data-ttu-id="34207-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="34207-131">TabOutlining</span></span>
- <span data-ttu-id="34207-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="34207-132">TabConflicts</span></span>
- <span data-ttu-id="34207-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="34207-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="34207-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="34207-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="34207-135">Excel</span><span class="sxs-lookup"><span data-stu-id="34207-135">Excel</span></span>

- <span data-ttu-id="34207-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="34207-136">**TabHome**</span></span>
- <span data-ttu-id="34207-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="34207-137">**TabInsert**</span></span>
- <span data-ttu-id="34207-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="34207-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="34207-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="34207-139">TabFormulas</span></span>
- <span data-ttu-id="34207-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="34207-140">**TabData**</span></span>
- <span data-ttu-id="34207-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="34207-141">**TabReview**</span></span>
- <span data-ttu-id="34207-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="34207-142">**TabView**</span></span>
- <span data-ttu-id="34207-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="34207-143">TabDeveloper</span></span>
- <span data-ttu-id="34207-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="34207-144">TabAddIns</span></span>
- <span data-ttu-id="34207-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="34207-145">TabPrintPreview</span></span>
- <span data-ttu-id="34207-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="34207-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="34207-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="34207-147">PowerPoint</span></span>

- <span data-ttu-id="34207-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="34207-148">**TabHome**</span></span>
- <span data-ttu-id="34207-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="34207-149">**TabInsert**</span></span>
- <span data-ttu-id="34207-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="34207-150">**TabDesign**</span></span>
- <span data-ttu-id="34207-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="34207-151">**TabTransitions**</span></span>
- <span data-ttu-id="34207-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="34207-152">**TabAnimations**</span></span>
- <span data-ttu-id="34207-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="34207-153">TabSlideShow</span></span>
- <span data-ttu-id="34207-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="34207-154">TabReview</span></span>
- <span data-ttu-id="34207-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="34207-155">**TabView**</span></span>
- <span data-ttu-id="34207-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="34207-156">TabDeveloper</span></span>
- <span data-ttu-id="34207-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="34207-157">TabAddIns</span></span>
- <span data-ttu-id="34207-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="34207-158">TabPrintPreview</span></span>
- <span data-ttu-id="34207-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="34207-159">TabMerge</span></span>
- <span data-ttu-id="34207-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="34207-160">TabGrayscale</span></span>
- <span data-ttu-id="34207-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="34207-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="34207-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="34207-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="34207-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="34207-163">TabSlideMaster</span></span>
- <span data-ttu-id="34207-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="34207-164">TabHandoutMaster</span></span>
- <span data-ttu-id="34207-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="34207-165">TabNotesMaster</span></span>
- <span data-ttu-id="34207-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="34207-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="34207-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="34207-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="34207-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="34207-168">OneNote</span></span>

- <span data-ttu-id="34207-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="34207-169">**TabHome**</span></span>
- <span data-ttu-id="34207-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="34207-170">**TabInsert**</span></span>
- <span data-ttu-id="34207-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="34207-171">**TabView**</span></span>
- <span data-ttu-id="34207-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="34207-172">TabDeveloper</span></span>
- <span data-ttu-id="34207-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="34207-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="34207-174">Group</span><span class="sxs-lookup"><span data-stu-id="34207-174">Group</span></span>

<span data-ttu-id="34207-p104">Groupe de points d’extension d’interface utilisateur dans un onglet. Un groupe peut contenir jusqu’à six contrôles. L’attribut **id** est requis et chaque **id** doit être unique au sein du manifeste. L’**ID** est une chaîne avec un maximum de 125 caractères. Voir l’[élément group](group.md).</span><span class="sxs-lookup"><span data-stu-id="34207-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="34207-179">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="34207-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
