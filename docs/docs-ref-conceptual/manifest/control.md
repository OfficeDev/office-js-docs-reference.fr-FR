# <a name="control-element"></a><span data-ttu-id="15a0b-101">Control, élément</span><span class="sxs-lookup"><span data-stu-id="15a0b-101">Control element</span></span>

<span data-ttu-id="15a0b-p101">Définit une fonction JavaScript qui exécute une action ou lance un volet Office. Un élément **Control** peut être une option de bouton ou de menu. Au moins un élément **Control** doit être inclus dans un élément [Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="15a0b-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="15a0b-105">Attributs</span><span class="sxs-lookup"><span data-stu-id="15a0b-105">Attributes</span></span>

|  <span data-ttu-id="15a0b-106">Attribut</span><span class="sxs-lookup"><span data-stu-id="15a0b-106">Attribute</span></span>  |  <span data-ttu-id="15a0b-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="15a0b-107">Required</span></span>  |  <span data-ttu-id="15a0b-108">Description</span><span class="sxs-lookup"><span data-stu-id="15a0b-108">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="15a0b-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="15a0b-109">**xsi:type**</span></span>|<span data-ttu-id="15a0b-110">Oui</span><span class="sxs-lookup"><span data-stu-id="15a0b-110">Yes</span></span>|<span data-ttu-id="15a0b-p102">Type de contrôle défini. Peut être `Button`, `Menu` ou `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="15a0b-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="15a0b-113">**id**</span><span class="sxs-lookup"><span data-stu-id="15a0b-113">**id**</span></span>|<span data-ttu-id="15a0b-114">Non</span><span class="sxs-lookup"><span data-stu-id="15a0b-114">No</span></span>|<span data-ttu-id="15a0b-p103">ID de l’élément de contrôle. Il doit comporter 125 caractères au maximum.</span><span class="sxs-lookup"><span data-stu-id="15a0b-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="15a0b-117">Le `MobileButton` valeur **xsi : type** est défini dans le schéma VersionOverrides 1.1.</span><span class="sxs-lookup"><span data-stu-id="15a0b-117">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="15a0b-118">Elle s’applique uniquement aux éléments de **contrôle de** contenu dans un élément [MobileFormFactor](mobileformfactor.md) .</span><span class="sxs-lookup"><span data-stu-id="15a0b-118">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="15a0b-119">Contrôle de bouton</span><span class="sxs-lookup"><span data-stu-id="15a0b-119">Button control</span></span>

<span data-ttu-id="15a0b-p105">Un bouton effectue une action unique quand il est sélectionné. Il peut exécuter une fonction ou afficher un volet Office. Chaque contrôle bouton doit avoir un `id` unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="15a0b-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="15a0b-123">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="15a0b-123">Child elements</span></span>
|  <span data-ttu-id="15a0b-124">Élément</span><span class="sxs-lookup"><span data-stu-id="15a0b-124">Element</span></span> |  <span data-ttu-id="15a0b-125">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="15a0b-125">Required</span></span>  |  <span data-ttu-id="15a0b-126">Description</span><span class="sxs-lookup"><span data-stu-id="15a0b-126">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="15a0b-127">**Étiquette**</span><span class="sxs-lookup"><span data-stu-id="15a0b-127">**Label**</span></span>     | <span data-ttu-id="15a0b-128">Oui</span><span class="sxs-lookup"><span data-stu-id="15a0b-128">Yes</span></span> |  <span data-ttu-id="15a0b-p106">Texte du bouton. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="15a0b-p106">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="15a0b-131">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="15a0b-131">**ToolTip**</span></span>  |<span data-ttu-id="15a0b-132">Non</span><span class="sxs-lookup"><span data-stu-id="15a0b-132">No</span></span>|<span data-ttu-id="15a0b-p107">Info-bulle pour le bouton. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String**. **String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="15a0b-p107">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="15a0b-136">Supertip</span><span class="sxs-lookup"><span data-stu-id="15a0b-136">Supertip</span></span>](supertip.md)  | <span data-ttu-id="15a0b-137">Oui</span><span class="sxs-lookup"><span data-stu-id="15a0b-137">Yes</span></span> |  <span data-ttu-id="15a0b-138">Info-bulle pour le bouton.</span><span class="sxs-lookup"><span data-stu-id="15a0b-138">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="15a0b-139">Icône</span><span class="sxs-lookup"><span data-stu-id="15a0b-139">Icon</span></span>](icon.md)      | <span data-ttu-id="15a0b-140">Oui</span><span class="sxs-lookup"><span data-stu-id="15a0b-140">Yes</span></span> |  <span data-ttu-id="15a0b-141">Image du bouton.</span><span class="sxs-lookup"><span data-stu-id="15a0b-141">An image for the button.</span></span>         |
|  [<span data-ttu-id="15a0b-142">Action</span><span class="sxs-lookup"><span data-stu-id="15a0b-142">Action</span></span>](action.md)    | <span data-ttu-id="15a0b-143">Oui</span><span class="sxs-lookup"><span data-stu-id="15a0b-143">Yes</span></span> |  <span data-ttu-id="15a0b-144">Spécifie l’action à effectuer.</span><span class="sxs-lookup"><span data-stu-id="15a0b-144">Specifies the action to perform.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="15a0b-145">Exemple du bouton ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="15a0b-145">ExecuteFunction button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="15a0b-146">Exemple du bouton ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="15a0b-146">ShowTaskpane button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="15a0b-147">Contrôles de menu (bouton déroulant)</span><span class="sxs-lookup"><span data-stu-id="15a0b-147">Menu (dropdown button) controls</span></span>

<span data-ttu-id="15a0b-p108">Un menu définit une liste statique d’options. Chaque option de menu exécute une fonction ou affiche un volet Office. Les sous-menus ne sont pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="15a0b-p108">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="15a0b-151">Lorsqu’il est utilisé avec un [point d’extension](extensionpoint.md) **PrimaryCommandSurface** ou **ContextMenu**, le contrôle de menu définit les éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="15a0b-151">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="15a0b-152">Un élément de menu au niveau racine.</span><span class="sxs-lookup"><span data-stu-id="15a0b-152">A root-level menu item.</span></span>

- <span data-ttu-id="15a0b-153">Une liste des éléments de sous-menu.</span><span class="sxs-lookup"><span data-stu-id="15a0b-153">A list of submenu items.</span></span>

<span data-ttu-id="15a0b-p109">Lorsqu’il est utilisé avec  **PrimaryCommandSurface**, l’élément de menu racine apparaît sous forme de bouton sur le ruban. Lorsque ce bouton est sélectionné, ce menu s’affiche comme une liste déroulante. Lorsqu’il est utilisé avec  **ContextMenu**, une option de menu comportant un sous-menu est inséré dans le menu contextuel. Dans les deux cas, les éléments de sous-menu individuels peuvent soit exécuter une fonction JavaScript, soit afficher un volet de tâches. Un seul niveau de sous-menus est actuellement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="15a0b-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="15a0b-p110">L’exemple suivant montre comment définir un élément de menu avec deux éléments de sous-menu. Le premier élément de sous-menu affiche un volet Office et le deuxième élément de sous-menu exécute une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="15a0b-p110">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

### <a name="child-elements"></a><span data-ttu-id="15a0b-161">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="15a0b-161">Child elements</span></span>

|  <span data-ttu-id="15a0b-162">Élément</span><span class="sxs-lookup"><span data-stu-id="15a0b-162">Element</span></span> |  <span data-ttu-id="15a0b-163">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="15a0b-163">Required</span></span>  |  <span data-ttu-id="15a0b-164">Description</span><span class="sxs-lookup"><span data-stu-id="15a0b-164">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="15a0b-165">**Étiquette**</span><span class="sxs-lookup"><span data-stu-id="15a0b-165">**Label**</span></span>     | <span data-ttu-id="15a0b-166">Oui</span><span class="sxs-lookup"><span data-stu-id="15a0b-166">Yes</span></span> |  <span data-ttu-id="15a0b-p111">Texte du bouton. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="15a0b-p111">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="15a0b-169">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="15a0b-169">**ToolTip**</span></span>  |<span data-ttu-id="15a0b-170">Non</span><span class="sxs-lookup"><span data-stu-id="15a0b-170">No</span></span>|<span data-ttu-id="15a0b-p112">Info-bulle pour le bouton. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String**. **String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="15a0b-p112">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="15a0b-174">Supertip</span><span class="sxs-lookup"><span data-stu-id="15a0b-174">Supertip</span></span>](supertip.md)  | <span data-ttu-id="15a0b-175">Oui</span><span class="sxs-lookup"><span data-stu-id="15a0b-175">Yes</span></span> |  <span data-ttu-id="15a0b-176">Info-bulle pour ce bouton.</span><span class="sxs-lookup"><span data-stu-id="15a0b-176">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="15a0b-177">Icône</span><span class="sxs-lookup"><span data-stu-id="15a0b-177">Icon</span></span>](icon.md)      | <span data-ttu-id="15a0b-178">Oui</span><span class="sxs-lookup"><span data-stu-id="15a0b-178">Yes</span></span> |  <span data-ttu-id="15a0b-179">Image du bouton.</span><span class="sxs-lookup"><span data-stu-id="15a0b-179">An image for the button.</span></span>         |
|  <span data-ttu-id="15a0b-180">**Items**</span><span class="sxs-lookup"><span data-stu-id="15a0b-180">**Items**</span></span>     | <span data-ttu-id="15a0b-181">Oui</span><span class="sxs-lookup"><span data-stu-id="15a0b-181">Yes</span></span> |  <span data-ttu-id="15a0b-p113">Ensemble de boutons à afficher dans le menu Contient les éléments **Item** pour chaque élément de sous-menu. Chaque élément **Item** contient les éléments enfants du [contrôle de bouton](#button-control).</span><span class="sxs-lookup"><span data-stu-id="15a0b-p113">A collection of Buttons to display within the menu. Contains the  **Item** elements for each submenu item. Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="15a0b-185">Exemples de contrôle de menu</span><span class="sxs-lookup"><span data-stu-id="15a0b-185">Menu control examples</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

## <a name="mobilebutton-control"></a><span data-ttu-id="15a0b-186">Contrôle MobileButton</span><span class="sxs-lookup"><span data-stu-id="15a0b-186">MobileButton control</span></span>

<span data-ttu-id="15a0b-p114">Un bouton mobile effectue une action unique lorsque l’utilisateur le sélectionne. Il peut exécuter une fonction ou afficher un volet Office. Chaque contrôle de bouton mobile doit avoir un `id` unique dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="15a0b-p114">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="15a0b-p115">La valeur `MobileButton` de **xsi:type** est définie dans le schéma VersionOverrides 1.1. Pour les éléments [VersionOverrides](versionoverrides.md) la contenant, l’attribut `xsi:type` doit avoir la valeur `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="15a0b-p115">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="15a0b-192">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="15a0b-192">Child elements</span></span>
|  <span data-ttu-id="15a0b-193">Élément</span><span class="sxs-lookup"><span data-stu-id="15a0b-193">Element</span></span> |  <span data-ttu-id="15a0b-194">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="15a0b-194">Required</span></span>  |  <span data-ttu-id="15a0b-195">Description</span><span class="sxs-lookup"><span data-stu-id="15a0b-195">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="15a0b-196">**Étiquette**</span><span class="sxs-lookup"><span data-stu-id="15a0b-196">**Label**</span></span>     | <span data-ttu-id="15a0b-197">Oui</span><span class="sxs-lookup"><span data-stu-id="15a0b-197">Yes</span></span> |  <span data-ttu-id="15a0b-p116">Texte du bouton. L’attribut **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="15a0b-p116">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="15a0b-200">Icône</span><span class="sxs-lookup"><span data-stu-id="15a0b-200">Icon</span></span>](icon.md)      | <span data-ttu-id="15a0b-201">Oui</span><span class="sxs-lookup"><span data-stu-id="15a0b-201">Yes</span></span> |  <span data-ttu-id="15a0b-202">Image du bouton.</span><span class="sxs-lookup"><span data-stu-id="15a0b-202">An image for the button.</span></span>         |
|  [<span data-ttu-id="15a0b-203">Action</span><span class="sxs-lookup"><span data-stu-id="15a0b-203">Action</span></span>](action.md)    | <span data-ttu-id="15a0b-204">Oui</span><span class="sxs-lookup"><span data-stu-id="15a0b-204">Yes</span></span> |  <span data-ttu-id="15a0b-205">Spécifie l’action à effectuer.</span><span class="sxs-lookup"><span data-stu-id="15a0b-205">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="15a0b-206">Exemple de bouton mobile ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="15a0b-206">ExecuteFunction mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="15a0b-207">Exemple de bouton mobile ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="15a0b-207">ShowTaskpane mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```