# <a name="versionoverrides-element"></a><span data-ttu-id="aa3db-101">Élément VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="aa3db-101">VersionOverrides element</span></span>

<span data-ttu-id="aa3db-p101">Élément racine qui contient des informations pour les commandes de complément implémentées par le complément. **VersionOverrides** est un élément enfant de l’élément [OfficeApp](./officeapp.md) dans le manifeste. Cet élément est pris en charge dans le schéma de manifeste v1.1 et versions ultérieures, mais est défini dans le schéma VersionOverrides v1.0 ou v1.1.</span><span class="sxs-lookup"><span data-stu-id="aa3db-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="aa3db-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="aa3db-105">Attributes</span></span>

|  <span data-ttu-id="aa3db-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="aa3db-106">Attribute</span></span>  |  <span data-ttu-id="aa3db-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="aa3db-107">Required</span></span>  |  <span data-ttu-id="aa3db-108">Description</span><span class="sxs-lookup"><span data-stu-id="aa3db-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="aa3db-109">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="aa3db-109">**xmlns**</span></span>       |  <span data-ttu-id="aa3db-110">Oui</span><span class="sxs-lookup"><span data-stu-id="aa3db-110">Yes</span></span>  |  <span data-ttu-id="aa3db-111">Emplacement du schéma, qui doit être `http://schemas.microsoft.com/office/mailappversionoverrides` lorsque `xsi:type` est `VersionOverridesV1_0`, et `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` lorsque `xsi:type` est `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="aa3db-111">The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.</span></span>|
|  <span data-ttu-id="aa3db-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="aa3db-112">**xsi:type**</span></span>  |  <span data-ttu-id="aa3db-113">Oui</span><span class="sxs-lookup"><span data-stu-id="aa3db-113">Yes</span></span>  | <span data-ttu-id="aa3db-p102">Version du schéma. À ce stade, les seules valeurs valides sont `VersionOverridesV1_0` et `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="aa3db-p102">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

> [!NOTE]
> <span data-ttu-id="aa3db-116">Actuellement uniquement 2016 Outlook prend en charge le schéma version 1.1 VersionOverrides et `VersionOverridesV1_1` type.</span><span class="sxs-lookup"><span data-stu-id="aa3db-116">Currently only Outlook 2016 supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="aa3db-117">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="aa3db-117">Child elements</span></span>

|  <span data-ttu-id="aa3db-118">Élément</span><span class="sxs-lookup"><span data-stu-id="aa3db-118">Element</span></span> |  <span data-ttu-id="aa3db-119">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="aa3db-119">Required</span></span>  |  <span data-ttu-id="aa3db-120">Description</span><span class="sxs-lookup"><span data-stu-id="aa3db-120">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="aa3db-121">**Description**</span><span class="sxs-lookup"><span data-stu-id="aa3db-121">**Description**</span></span>    |  <span data-ttu-id="aa3db-122">Non</span><span class="sxs-lookup"><span data-stu-id="aa3db-122">No</span></span>   |  <span data-ttu-id="aa3db-p103">Décrit le complément. Cela remplace l’élément `Description` dans une partie parent du manifeste. Le texte de la description est contenu dans un élément enfant de l’élément **LongString** contenu dans l’élément [Resources](./resources.md). L’attribut `resid` de l’élément **Description** est défini sur la valeur de l’attribut `id` de l’élément `String` qui contient le texte.</span><span class="sxs-lookup"><span data-stu-id="aa3db-p103">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="aa3db-127">**Configuration requise**</span><span class="sxs-lookup"><span data-stu-id="aa3db-127">**Requirements**</span></span>  |  <span data-ttu-id="aa3db-128">Non</span><span class="sxs-lookup"><span data-stu-id="aa3db-128">No</span></span>   |  <span data-ttu-id="aa3db-p104">Spécifie l’ensemble de conditions requises minimal et la version d’Office.js qui doit être activée par le complément Office. Cela remplace l’élément `Requirements` dans la partie parent du manifeste.</span><span class="sxs-lookup"><span data-stu-id="aa3db-p104">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="aa3db-131">Hôtes</span><span class="sxs-lookup"><span data-stu-id="aa3db-131">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="aa3db-132">Oui</span><span class="sxs-lookup"><span data-stu-id="aa3db-132">Yes</span></span>  |  <span data-ttu-id="aa3db-p105">Spécifie une collection d’hôtes d’Office. L’élément Hosts enfant remplace l’élément Hosts dans la partie parent du manifeste.</span><span class="sxs-lookup"><span data-stu-id="aa3db-p105">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="aa3db-135">Ressources</span><span class="sxs-lookup"><span data-stu-id="aa3db-135">Resources</span></span>](./resources.md)    |  <span data-ttu-id="aa3db-136">Oui</span><span class="sxs-lookup"><span data-stu-id="aa3db-136">Yes</span></span>  | <span data-ttu-id="aa3db-137">Définit une collection de ressources (chaînes, URL et images) qui sont référencées par d’autres éléments de manifeste.</span><span class="sxs-lookup"><span data-stu-id="aa3db-137">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  <span data-ttu-id="aa3db-138">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="aa3db-138">**VersionOverrides**</span></span>    |  <span data-ttu-id="aa3db-139">Non</span><span class="sxs-lookup"><span data-stu-id="aa3db-139">No</span></span>  | <span data-ttu-id="aa3db-p106">Définit des commandes de complément sous une version plus récente du schéma. Voir [Mise en œuvre de plusieurs versions](#implementing-multiple-versions) pour plus d’informations.</span><span class="sxs-lookup"><span data-stu-id="aa3db-p106">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  <span data-ttu-id="aa3db-142">**WebApplicationInfo**</span><span class="sxs-lookup"><span data-stu-id="aa3db-142">**WebApplicationInfo**</span></span>    |  <span data-ttu-id="aa3db-143">Non</span><span class="sxs-lookup"><span data-stu-id="aa3db-143">No</span></span>  | <span data-ttu-id="aa3db-144">Fournit des détails sur l’application web associée au complément.</span><span class="sxs-lookup"><span data-stu-id="aa3db-144">Specifies details about the add-in's associated Web application.</span></span> |



### <a name="versionoverrides-example"></a><span data-ttu-id="aa3db-145">Exemple VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="aa3db-145">VersionOverrides example</span></span>
```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a><span data-ttu-id="aa3db-146">Mise en œuvre de plusieurs versions</span><span class="sxs-lookup"><span data-stu-id="aa3db-146">Implementing multiple versions</span></span>

<span data-ttu-id="aa3db-p107">Un manifeste peut implémenter plusieurs versions de l’élément `VersionOverrides` qui prennent en charge différentes versions du schéma VersionOverrides. Cette opération permet éventuellement la prise en charge de nouvelles fonctionnalités dans un schéma plus récent tout en prenant en charge des clients plus anciens qui ne prennent pas en charge les nouvelles fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="aa3db-p107">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="aa3db-149">Pour mettre en œuvre plusieurs versions, l’élément `VersionOverrides` de la nouvelle version doit être un enfant de l’élément `VersionOverrides` de l’ancienne version.</span><span class="sxs-lookup"><span data-stu-id="aa3db-149">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="aa3db-150">L’élément enfant `VersionOverrides` n’hérite pas des valeurs du parent.</span><span class="sxs-lookup"><span data-stu-id="aa3db-150">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="aa3db-151">Pour mettre en œuvre à la fois les schémas VersionOverrides v1.0 et v1.1, le manifeste devrait ressembler à l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="aa3db-151">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
...
</OfficeApp>
```
