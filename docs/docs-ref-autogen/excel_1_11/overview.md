---
title: Référence de l’API JavaScript d’Office
description: Les conditions requises de l’API JavaScript pour Office par hôte.
ms.date: 05/05/2020
ms.openlocfilehash: 3a32c47b23fd6635c4c2b44b58ee9b351fffd8d5
ms.sourcegitcommit: c38b0b4cf446e10aed0f7de4401cf7060020d260
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2020
ms.locfileid: "44114005"
---
# <a name="office-javascript-api-reference"></a><span data-ttu-id="4ce4b-103">Référence de l’API JavaScript d’Office</span><span class="sxs-lookup"><span data-stu-id="4ce4b-103">Office JavaScript API reference</span></span>

<span data-ttu-id="4ce4b-104">L’interface API JavaScript pour Office vous permet de créer des applications web qui interagissent avec les modèles objet dans les applications hôtes Office.</span><span class="sxs-lookup"><span data-stu-id="4ce4b-104">The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications.</span></span> <span data-ttu-id="4ce4b-105">Utilisez cette section pour en savoir plus sur les classes, les méthodes et les autres types disponibles pour créer des compléments Office.</span><span class="sxs-lookup"><span data-stu-id="4ce4b-105">Use this section to learn more about the classes, methods, and other types available for building Office Add-ins.</span></span>

<span data-ttu-id="4ce4b-106">La liste suivante répertorie les ensembles de conditions requises propres à l’hôte (et les API communes entre hôtes).</span><span class="sxs-lookup"><span data-stu-id="4ce4b-106">The following is a list of host-specific requirement sets (and the cross-host Common APIs).</span></span> <span data-ttu-id="4ce4b-107">Chaque élément est lié à une version de la documentation de référence de l’API prise en charge par cet ensemble de conditions (par exemple, ExcelApi 1,3 affiche les API dans ExcelApi 1,1, 1,2, 1,3, ainsi que l’API commune).</span><span class="sxs-lookup"><span data-stu-id="4ce4b-107">Each item links to a version of the API reference documentation that is supported by that requirement set (e.g. ExcelApi 1.3 shows APIs in ExcelApi 1.1, 1.2, 1.3 as well as the Common API).</span></span>

<span data-ttu-id="4ce4b-108">`ExcelApiOnline 1.1`est un ensemble de conditions requises spéciales.</span><span class="sxs-lookup"><span data-stu-id="4ce4b-108">`ExcelApiOnline 1.1` is a special requirement set.</span></span> <span data-ttu-id="4ce4b-109">Il contient les dernières API pour Excel sur le Web, mais ces API ne sont peut-être pas encore entièrement prises en charge sur toutes les plateformes.</span><span class="sxs-lookup"><span data-stu-id="4ce4b-109">It contains the latest APIs for Excel on the web, but those APIs may not yet be fully supported across all platforms.</span></span> <span data-ttu-id="4ce4b-110">Pour plus d’informations, consultez l' [ensemble de conditions requises pour l’API JavaScript pour Excel Online uniquement](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set) .</span><span class="sxs-lookup"><span data-stu-id="4ce4b-110">See [Excel JavaScript API online-only requirement set](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set) for more information.</span></span>

> [!TIP]
> <span data-ttu-id="4ce4b-111">Choisissez un lien dans cette page pour afficher la documentation de référence pour les API prises en charge par l’ensemble de conditions spécifié, ou utilisez le menu déroulant sélection de filtre au-dessus de la table des matières pour modifier l’ensemble de conditions requises à tout moment.</span><span class="sxs-lookup"><span data-stu-id="4ce4b-111">Choose a link on this page to view reference documentation for APIs supported by the specified requirement set, or use the filter selection drop-down menu above the table of contents to change the requirement set at any time.</span></span>

## <a name="excel"></a><span data-ttu-id="4ce4b-112">Excel</span><span class="sxs-lookup"><span data-stu-id="4ce4b-112">Excel</span></span>

- [<span data-ttu-id="4ce4b-113">Préversion ExcelApi</span><span class="sxs-lookup"><span data-stu-id="4ce4b-113">ExcelApi Preview</span></span>](/javascript/api/excel?view=excel-js-preview)
- [<span data-ttu-id="4ce4b-114">ExcelApiOnline 1,1</span><span class="sxs-lookup"><span data-stu-id="4ce4b-114">ExcelApiOnline 1.1</span></span>](/javascript/api/excel?view=excel-js-online)
- [<span data-ttu-id="4ce4b-115">ExcelApi 1,11</span><span class="sxs-lookup"><span data-stu-id="4ce4b-115">ExcelApi 1.11</span></span>](/javascript/api/excel?view=excel-js-1.11)
- [<span data-ttu-id="4ce4b-116">ExcelApi 1.10</span><span class="sxs-lookup"><span data-stu-id="4ce4b-116">ExcelApi 1.10</span></span>](/javascript/api/excel?view=excel-js-1.10)
- [<span data-ttu-id="4ce4b-117">ExcelApi 1.9</span><span class="sxs-lookup"><span data-stu-id="4ce4b-117">ExcelApi 1.9</span></span>](/javascript/api/excel?view=excel-js-1.9)
- [<span data-ttu-id="4ce4b-118">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="4ce4b-118">ExcelApi 1.8</span></span>](/javascript/api/excel?view=excel-js-1.8)
- [<span data-ttu-id="4ce4b-119">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="4ce4b-119">ExcelApi 1.7</span></span>](/javascript/api/excel?view=excel-js-1.7)
- [<span data-ttu-id="4ce4b-120">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="4ce4b-120">ExcelApi 1.6</span></span>](/javascript/api/excel?view=excel-js-1.6)
- [<span data-ttu-id="4ce4b-121">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="4ce4b-121">ExcelApi 1.5</span></span>](/javascript/api/excel?view=excel-js-1.5)
- [<span data-ttu-id="4ce4b-122">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="4ce4b-122">ExcelApi 1.4</span></span>](/javascript/api/excel?view=excel-js-1.4)
- [<span data-ttu-id="4ce4b-123">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="4ce4b-123">ExcelApi 1.3</span></span>](/javascript/api/excel?view=excel-js-1.3)
- [<span data-ttu-id="4ce4b-124">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="4ce4b-124">ExcelApi 1.2</span></span>](/javascript/api/excel?view=excel-js-1.2)
- [<span data-ttu-id="4ce4b-125">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="4ce4b-125">ExcelApi 1.1</span></span>](/javascript/api/excel?view=excel-js-1.1)

## <a name="onenote"></a><span data-ttu-id="4ce4b-126">OneNote</span><span class="sxs-lookup"><span data-stu-id="4ce4b-126">OneNote</span></span>

- [<span data-ttu-id="4ce4b-127">OneNote 1,1</span><span class="sxs-lookup"><span data-stu-id="4ce4b-127">OneNote 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)

## <a name="outlook"></a><span data-ttu-id="4ce4b-128">Outlook</span><span class="sxs-lookup"><span data-stu-id="4ce4b-128">Outlook</span></span>

- [<span data-ttu-id="4ce4b-129">Aperçu de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ce4b-129">Mailbox Preview</span></span>](/javascript/api/outlook?view=outlook-js-preview)
- [<span data-ttu-id="4ce4b-130">Mailbox 1.8</span><span class="sxs-lookup"><span data-stu-id="4ce4b-130">Mailbox 1.8</span></span>](/javascript/api/outlook?view=outlook-js-1.8)
- [<span data-ttu-id="4ce4b-131">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="4ce4b-131">Mailbox 1.7</span></span>](/javascript/api/outlook?view=outlook-js-1.7)
- [<span data-ttu-id="4ce4b-132">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="4ce4b-132">Mailbox 1.6</span></span>](/javascript/api/outlook?view=outlook-js-1.6)
- [<span data-ttu-id="4ce4b-133">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="4ce4b-133">Mailbox 1.5</span></span>](/javascript/api/outlook?view=outlook-js-1.5)
- [<span data-ttu-id="4ce4b-134">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="4ce4b-134">Mailbox 1.4</span></span>](/javascript/api/outlook?view=outlook-js-1.4)
- [<span data-ttu-id="4ce4b-135">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="4ce4b-135">Mailbox 1.3</span></span>](/javascript/api/outlook?view=outlook-js-1.3)
- [<span data-ttu-id="4ce4b-136">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="4ce4b-136">Mailbox 1.2</span></span>](/javascript/api/outlook?view=outlook-js-1.2)
- [<span data-ttu-id="4ce4b-137">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="4ce4b-137">Mailbox 1.1</span></span>](/javascript/api/outlook?view=outlook-js-1.1)

## <a name="powerpoint"></a><span data-ttu-id="4ce4b-138">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4ce4b-138">PowerPoint</span></span>

- [<span data-ttu-id="4ce4b-139">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="4ce4b-139">PowerPointApi 1.1</span></span>](/javascript/api/powerpoint?view=powerpoint-js-1.1)

## <a name="visio"></a><span data-ttu-id="4ce4b-140">Visio</span><span class="sxs-lookup"><span data-stu-id="4ce4b-140">Visio</span></span>

- [<span data-ttu-id="4ce4b-141">VisioApi 1,1</span><span class="sxs-lookup"><span data-stu-id="4ce4b-141">VisioApi 1.1</span></span>](/javascript/api/visio?view=visio-js-1.1)

## <a name="word"></a><span data-ttu-id="4ce4b-142">Word</span><span class="sxs-lookup"><span data-stu-id="4ce4b-142">Word</span></span>

- [<span data-ttu-id="4ce4b-143">Aperçu de Word</span><span class="sxs-lookup"><span data-stu-id="4ce4b-143">Word Preview</span></span>](/javascript/api/word?view=word-js-preview)
- [<span data-ttu-id="4ce4b-144">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="4ce4b-144">WordApi 1.3</span></span>](/javascript/api/word?view=word-js-1.3)
- [<span data-ttu-id="4ce4b-145">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="4ce4b-145">WordApi 1.2</span></span>](/javascript/api/word?view=word-js-1.2)
- [<span data-ttu-id="4ce4b-146">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="4ce4b-146">WordApi 1.1</span></span>](/javascript/api/word?view=word-js-1.1)

## <a name="common-api"></a><span data-ttu-id="4ce4b-147">API communes</span><span class="sxs-lookup"><span data-stu-id="4ce4b-147">Common API</span></span>

- [<span data-ttu-id="4ce4b-148">API communes</span><span class="sxs-lookup"><span data-stu-id="4ce4b-148">Common API</span></span>](/javascript/api/office?view=common-js)

## <a name="see-also"></a><span data-ttu-id="4ce4b-149">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4ce4b-149">See also</span></span>

- [<span data-ttu-id="4ce4b-150">À propos des compléments Office</span><span class="sxs-lookup"><span data-stu-id="4ce4b-150">About Office Add-ins</span></span>](/office/dev/add-ins/overview)
- [<span data-ttu-id="4ce4b-151">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="4ce4b-151">Office Add-in host and platform availability</span></span>](/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="4ce4b-152">Versions d’Office et ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ce4b-152">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="4ce4b-153">Explorer l’API JavaScript Office à l’aide de Script Lab</span><span class="sxs-lookup"><span data-stu-id="4ce4b-153">Explore Office JavaScript API using Script Lab</span></span>](/office/dev/add-ins/overview/explore-with-script-lab)