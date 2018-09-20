# <a name="extensionpoint-element"></a><span data-ttu-id="68fa4-101">Élément ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="68fa4-101">ExtensionPoint element</span></span>

 <span data-ttu-id="68fa4-102">Définit l’emplacement où se trouvent les fonctionnalités d’un complément dans l’interface utilisateur Office.</span><span class="sxs-lookup"><span data-stu-id="68fa4-102">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="68fa4-103">L’élément **ExtensionPoint** est un élément enfant de [AllFormFactors](allformfactors.md) ou [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="68fa4-103">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="68fa4-104">Attributs</span><span class="sxs-lookup"><span data-stu-id="68fa4-104">Attributes</span></span>

|  <span data-ttu-id="68fa4-105">Attribut</span><span class="sxs-lookup"><span data-stu-id="68fa4-105">Attribute</span></span>  |  <span data-ttu-id="68fa4-106">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="68fa4-106">Required</span></span>  |  <span data-ttu-id="68fa4-107">Description</span><span class="sxs-lookup"><span data-stu-id="68fa4-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="68fa4-108">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="68fa4-108">**xsi:type**</span></span>  |  <span data-ttu-id="68fa4-109">Oui</span><span class="sxs-lookup"><span data-stu-id="68fa4-109">Yes</span></span>  | <span data-ttu-id="68fa4-110">Type de point d’extension défini.</span><span class="sxs-lookup"><span data-stu-id="68fa4-110">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="68fa4-111">Points d’extension pour Excel uniquement</span><span class="sxs-lookup"><span data-stu-id="68fa4-111">Extension points for Excel only</span></span>

- <span data-ttu-id="68fa4-112">**CustomFunctions** - une fonction personnalisée écrite dans JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="68fa4-112">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="68fa4-113">[Exemple de code XML ce](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) montre comment utiliser l’élément **ExtensionPoint** avec la valeur d’attribut **CustomFunctions** et les éléments enfants à utiliser.</span><span class="sxs-lookup"><span data-stu-id="68fa4-113">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="68fa4-114">Points d’extension pour les commandes de complément Word, Excel, PowerPoint et OneNote</span><span class="sxs-lookup"><span data-stu-id="68fa4-114">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="68fa4-115">**PrimaryCommandSurface** : ruban dans Office.</span><span class="sxs-lookup"><span data-stu-id="68fa4-115">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="68fa4-116">**ContextMenu** : menu contextuel qui apparaît lorsque vous cliquez avec le bouton droit de la souris dans l’interface utilisateur Office.</span><span class="sxs-lookup"><span data-stu-id="68fa4-116">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="68fa4-117">Les exemples suivants montrent comment utiliser l’élément  **ExtensionPoint** avec les valeurs d’attribut **PrimaryCommandSurface** et **ContextMenu**, ainsi que les éléments enfants qui doivent être utilisés avec chacune d’elles.</span><span class="sxs-lookup"><span data-stu-id="68fa4-117">The following examples show how to use the  **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="68fa4-118">Pour les éléments qui contiennent un attribut ID, vérifiez que vous fournissez un ID unique.</span><span class="sxs-lookup"><span data-stu-id="68fa4-118">For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="68fa4-119">Nous recommandons d’utiliser le nom de votre société, ainsi que votre identifiant.</span><span class="sxs-lookup"><span data-stu-id="68fa4-119">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="68fa4-120">Par exemple, utilisez la syntaxe suivante.</span><span class="sxs-lookup"><span data-stu-id="68fa4-120">For example, use the following format.</span></span> <CustomTab id="mycompanyname.mygroupname">

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
          <CustomTab id="Contoso Tab">
          <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
            <!-- <OfficeTab id="TabData"> -->
            <Label resid="residLabel4" />
            <Group id="Group1Id12">
              <Label resid="residLabel4" />
              <Icon>
                <bt:Image size="16" resid="icon1_32x32" />
                <bt:Image size="32" resid="icon1_32x32" />
                <bt:Image size="80" resid="icon1_32x32" />
              </Icon>
              <Tooltip resid="residToolTip" />
              <Control xsi:type="Button" id="Button1Id1">

                  <!-- information about the control -->
              </Control>
              <!-- other controls, as needed -->
            </Group>
          </CustomTab>
        </ExtensionPoint>

      <ExtensionPoint xsi:type="ContextMenu">
        <OfficeMenu id="ContextMenuCell">
          <Control xsi:type="Menu" id="ContextMenu2">
                  <!-- information about the control -->
          </Control>
          <!-- other controls, as needed -->
        </OfficeMenu>
        </ExtensionPoint>
```

#### <a name="child-elements"></a><span data-ttu-id="68fa4-121">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="68fa4-121">Child elements</span></span>
 
|<span data-ttu-id="68fa4-122">**Élément**</span><span class="sxs-lookup"><span data-stu-id="68fa4-122">**Element**</span></span>|<span data-ttu-id="68fa4-123">**Description**</span><span class="sxs-lookup"><span data-stu-id="68fa4-123">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="68fa4-124">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="68fa4-124">**CustomTab**</span></span>|<span data-ttu-id="68fa4-p103">Obligatoire pour ajouter un onglet personnalisé au ruban (en utilisant  **PrimaryCommandSurface**). Si vous utilisez l’élément  **CustomTab**, vous ne pouvez pas utiliser l’élément  **OfficeTab**. L’attribut  **id** est requis.</span><span class="sxs-lookup"><span data-stu-id="68fa4-p103">Required if you want to add a custom tab to the ribbon (using  **PrimaryCommandSurface**). If you use the  **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="68fa4-128">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="68fa4-128">**OfficeTab**</span></span>|<span data-ttu-id="68fa4-p104">Obligatoire pour étendre un onglet du ruban Office par défaut (en utilisant **PrimaryCommandSurface**). Si vous utilisez l’élément **OfficeTab**, vous ne pouvez pas utiliser l’élément **CustomTab**. Pour plus d’informations, voir [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="68fa4-p104">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the  **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="68fa4-132">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="68fa4-132">**OfficeMenu**</span></span>|<span data-ttu-id="68fa4-p105">Obligatoire pour ajouter des commandes de complément à un menu contextuel par défaut (en utilisant **ContextMenu**). L’attribut **id** doit être défini sur : </span><span class="sxs-lookup"><span data-stu-id="68fa4-p105">Required if you're adding add-in commands to a default context menu (using  **ContextMenu**). The  **id** attribute must be set to: </span></span><br/> <span data-ttu-id="68fa4-p106">- **ContextMenuText** pour Excel ou Word. Affiche l’élément dans le menu contextuel lorsque du texte est sélectionné et que l’utilisateur clique dessus avec le bouton droit de la souris. </span><span class="sxs-lookup"><span data-stu-id="68fa4-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="68fa4-p107">- **ContextMenuCell** pour Excel. Affiche l’élément dans le menu contextuel lorsque l’utilisateur clique avec le bouton droit de la souris dans une cellule de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="68fa4-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="68fa4-139">**Group**</span><span class="sxs-lookup"><span data-stu-id="68fa4-139">**Group**</span></span>|<span data-ttu-id="68fa4-p108">Groupe de points d’extension de l’interface utilisateur sur un onglet. Un groupe peut comporter jusqu’à six contrôles. L’attribut  **id** est requis. Il s’agit d’une chaîne contenant un maximum de 125 caractères.</span><span class="sxs-lookup"><span data-stu-id="68fa4-p108">A group of user interface extension points on a tab. A group can have up to six controls. The  **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="68fa4-143">**Label**</span><span class="sxs-lookup"><span data-stu-id="68fa4-143">**Label**</span></span>|<span data-ttu-id="68fa4-p109">Obligatoire. Libellé du groupe. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String**. L’élément  **String** est un enfant de l’élément **ShortStrings**, qui est lui-même un enfant de l’élément  **Resources**.</span><span class="sxs-lookup"><span data-stu-id="68fa4-p109">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="68fa4-148">**Icon**</span><span class="sxs-lookup"><span data-stu-id="68fa4-148">**Icon**</span></span>|<span data-ttu-id="68fa4-p110">Obligatoire. Indique l’icône du groupe qui doit être utilisée sur les périphériques de petit facteur de forme ou lorsque les boutons sont affichés en trop grand nombre. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **Image**. L’élément  **Image** est un enfant de l’élément **Images**, qui est lui-même un enfant de l’élément  **Resources**. L’attribut **size** donne la taille, en pixels, de l’image. Trois tailles d’image, en pixels, sont obligatoires : 16, 32 et 80. Cinq tailles facultatives, en pixels, sont également prises en charge : 20, 24, 40, 48 et 64.</span><span class="sxs-lookup"><span data-stu-id="68fa4-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="68fa4-156">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="68fa4-156">**Tooltip**</span></span>|<span data-ttu-id="68fa4-p111">Facultatif. Info-bulle du groupe. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String**. L’élément  **String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément  **Resources**.</span><span class="sxs-lookup"><span data-stu-id="68fa4-p111">Optional. The tooltip of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="68fa4-161">**Control**</span><span class="sxs-lookup"><span data-stu-id="68fa4-161">**Control**</span></span>|<span data-ttu-id="68fa4-162">Chaque groupe requiert au moins un contrôle.</span><span class="sxs-lookup"><span data-stu-id="68fa4-162">Each group requires at least one control.</span></span> <span data-ttu-id="68fa4-163">Un élément **Control** peut être un **bouton** ou un **menu**.</span><span class="sxs-lookup"><span data-stu-id="68fa4-163">A  **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="68fa4-164">Utilisez un **menu** pour spécifier une liste déroulante de contrôles de bouton.</span><span class="sxs-lookup"><span data-stu-id="68fa4-164">Use  **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="68fa4-165">Actuellement, seuls les boutons et les menus sont pris en charge.</span><span class="sxs-lookup"><span data-stu-id="68fa4-165">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="68fa4-166">Pour plus d’informations, reportez-vous aux sections [Contrôles de bouton](control.md#button-control) et [Contrôles de menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="68fa4-166">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="68fa4-167">**Remarque :**  Pour faciliter le dépannage, nous recommandons d’un élément de **contrôle** et les éléments enfants de **ressources** connexes ajouter un à la fois.</span><span class="sxs-lookup"><span data-stu-id="68fa4-167">**Note:**  To make troubleshooting easier, we recommend that a  **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="68fa4-168">**Script**</span><span class="sxs-lookup"><span data-stu-id="68fa4-168">**Script**</span></span>|<span data-ttu-id="68fa4-169">Liens vers le fichier JavaScript avec la définition de la fonction personnalisée et le code d’inscription.</span><span class="sxs-lookup"><span data-stu-id="68fa4-169">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="68fa4-170">Cet élément n’est pas utilisé dans l’aperçu pour les développeurs.</span><span class="sxs-lookup"><span data-stu-id="68fa4-170">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="68fa4-171">À la place, la page HTML est responsable du chargement de tous les fichiers JavaScript.</span><span class="sxs-lookup"><span data-stu-id="68fa4-171">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="68fa4-172">**Page**</span><span class="sxs-lookup"><span data-stu-id="68fa4-172">**Page**</span></span>|<span data-ttu-id="68fa4-173">Liens vers la page HTML de vos fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="68fa4-173">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="68fa4-174">Points d’extension pour Outlook</span><span class="sxs-lookup"><span data-stu-id="68fa4-174">Extension points for Outlook</span></span>

- [<span data-ttu-id="68fa4-175">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="68fa4-175">MessageReadCommandSurface</span></span>](#messagereadcommandsurface) 
- [<span data-ttu-id="68fa4-176">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="68fa4-176">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface) 
- [<span data-ttu-id="68fa4-177">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="68fa4-177">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface) 
- [<span data-ttu-id="68fa4-178">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="68fa4-178">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="68fa4-179">[Module](#module) (peut uniquement être utilisé dans [DesktopFormFactor](desktopformfactor.md).)</span><span class="sxs-lookup"><span data-stu-id="68fa4-179">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="68fa4-180">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="68fa4-180">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="68fa4-181">Événements</span><span class="sxs-lookup"><span data-stu-id="68fa4-181">Events</span></span>](#events)
- [<span data-ttu-id="68fa4-182">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="68fa4-182">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="68fa4-183">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="68fa4-183">MessageReadCommandSurface</span></span>
<span data-ttu-id="68fa4-p114">Ce point d’extension place des boutons dans la surface de commande pour le mode de lecture de courrier électronique. Dans l’application de bureau Outlook, cela apparaît dans le ruban.</span><span class="sxs-lookup"><span data-stu-id="68fa4-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="68fa4-186">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="68fa4-186">Child elements</span></span>

|  <span data-ttu-id="68fa4-187">Élément</span><span class="sxs-lookup"><span data-stu-id="68fa4-187">Element</span></span> |  <span data-ttu-id="68fa4-188">Description</span><span class="sxs-lookup"><span data-stu-id="68fa4-188">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="68fa4-189">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-189">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="68fa4-190">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="68fa4-190">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="68fa4-191">CustomTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-191">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="68fa4-192">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="68fa4-192">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="68fa4-193">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-193">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="68fa4-194">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-194">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="68fa4-195">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="68fa4-195">MessageComposeCommandSurface</span></span>
<span data-ttu-id="68fa4-196">Ce point d’extension place des boutons sur le ruban pour les compléments à l’aide du formulaire de composition de messagerie.</span><span class="sxs-lookup"><span data-stu-id="68fa4-196">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="68fa4-197">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="68fa4-197">Child elements</span></span>

|  <span data-ttu-id="68fa4-198">Élément</span><span class="sxs-lookup"><span data-stu-id="68fa4-198">Element</span></span> |  <span data-ttu-id="68fa4-199">Description</span><span class="sxs-lookup"><span data-stu-id="68fa4-199">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="68fa4-200">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-200">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="68fa4-201">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="68fa4-201">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="68fa4-202">CustomTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-202">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="68fa4-203">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="68fa4-203">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="68fa4-204">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-204">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="68fa4-205">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-205">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="68fa4-206">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="68fa4-206">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="68fa4-207">Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention de l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="68fa4-207">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="68fa4-208">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="68fa4-208">Child elements</span></span>

|  <span data-ttu-id="68fa4-209">Élément</span><span class="sxs-lookup"><span data-stu-id="68fa4-209">Element</span></span> |  <span data-ttu-id="68fa4-210">Description</span><span class="sxs-lookup"><span data-stu-id="68fa4-210">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="68fa4-211">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-211">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="68fa4-212">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="68fa4-212">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="68fa4-213">CustomTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-213">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="68fa4-214">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="68fa4-214">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="68fa4-215">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-215">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="68fa4-216">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-216">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="68fa4-217">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="68fa4-217">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="68fa4-218">Ce point d’extension place des boutons sur le ruban pour le formulaire qui est affiché à l’intention du participant à la réunion.</span><span class="sxs-lookup"><span data-stu-id="68fa4-218">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="68fa4-219">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="68fa4-219">Child elements</span></span>

|  <span data-ttu-id="68fa4-220">Élément</span><span class="sxs-lookup"><span data-stu-id="68fa4-220">Element</span></span> |  <span data-ttu-id="68fa4-221">Description</span><span class="sxs-lookup"><span data-stu-id="68fa4-221">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="68fa4-222">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-222">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="68fa4-223">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="68fa4-223">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="68fa4-224">CustomTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-224">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="68fa4-225">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="68fa4-225">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="68fa4-226">Exemple OfficeTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-226">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="68fa4-227">Exemple CustomTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-227">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="68fa4-228">Module</span><span class="sxs-lookup"><span data-stu-id="68fa4-228">Module</span></span>

<span data-ttu-id="68fa4-229">Ce point d’extension place des boutons sur le ruban pour l’extension de module.</span><span class="sxs-lookup"><span data-stu-id="68fa4-229">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="68fa4-230">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="68fa4-230">Child elements</span></span>

|  <span data-ttu-id="68fa4-231">Élément</span><span class="sxs-lookup"><span data-stu-id="68fa4-231">Element</span></span> |  <span data-ttu-id="68fa4-232">Description</span><span class="sxs-lookup"><span data-stu-id="68fa4-232">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="68fa4-233">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-233">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="68fa4-234">Ajoute les commandes à l’onglet de ruban par défaut.</span><span class="sxs-lookup"><span data-stu-id="68fa4-234">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="68fa4-235">CustomTab</span><span class="sxs-lookup"><span data-stu-id="68fa4-235">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="68fa4-236">Ajoute les commandes à l’onglet de ruban personnalisé.</span><span class="sxs-lookup"><span data-stu-id="68fa4-236">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="68fa4-237">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="68fa4-237">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="68fa4-238">Ce point d’extension place des boutons dans la surface de commande pour le mode de lecture de courrier électronique dans le facteur de forme pour environnement mobile.</span><span class="sxs-lookup"><span data-stu-id="68fa4-238">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="68fa4-239">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="68fa4-239">Child elements</span></span>

|  <span data-ttu-id="68fa4-240">Élément</span><span class="sxs-lookup"><span data-stu-id="68fa4-240">Element</span></span> |  <span data-ttu-id="68fa4-241">Description</span><span class="sxs-lookup"><span data-stu-id="68fa4-241">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="68fa4-242">Group</span><span class="sxs-lookup"><span data-stu-id="68fa4-242">Group</span></span>](group.md) |  <span data-ttu-id="68fa4-243">Ajoute un groupe de boutons à la surface de commande.</span><span class="sxs-lookup"><span data-stu-id="68fa4-243">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="68fa4-244">Les éléments **ExtensionPoint** de ce type peuvent uniquement avoir un élément enfant, à savoir un élément **Group**.</span><span class="sxs-lookup"><span data-stu-id="68fa4-244">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="68fa4-245">Pour les éléments **Control** contenus dans ce point d’extension, l’attribut **xsi:type** doit avoir la valeur `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="68fa4-245">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="68fa4-246">Exemple</span><span class="sxs-lookup"><span data-stu-id="68fa4-246">Example</span></span>
```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
      <Control id="mobileButton1" xsi:type="MobileButton">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="events"></a><span data-ttu-id="68fa4-247">Événements</span><span class="sxs-lookup"><span data-stu-id="68fa4-247">Events</span></span>

<span data-ttu-id="68fa4-248">Ce point d’extension ajoute un gestionnaire d’événements pour un événement spécifié.</span><span class="sxs-lookup"><span data-stu-id="68fa4-248">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="68fa4-249">Ce type d’élément est uniquement pris en charge par Outlook sur le web dans Office 365.</span><span class="sxs-lookup"><span data-stu-id="68fa4-249">This element type is only supported by Outlook on the web in Office 365.</span></span>

| <span data-ttu-id="68fa4-250">Élément</span><span class="sxs-lookup"><span data-stu-id="68fa4-250">Element</span></span> | <span data-ttu-id="68fa4-251">Description</span><span class="sxs-lookup"><span data-stu-id="68fa4-251">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="68fa4-252">Event</span><span class="sxs-lookup"><span data-stu-id="68fa4-252">Event</span></span>](event.md) |  <span data-ttu-id="68fa4-253">Indique l’événement et la fonction gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="68fa4-253">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="68fa4-254">Exemple d’événement ItemSend</span><span class="sxs-lookup"><span data-stu-id="68fa4-254">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events"> 
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
</ExtensionPoint> 
```

### <a name="detectedentity"></a><span data-ttu-id="68fa4-255">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="68fa4-255">DetectedEntity</span></span>

<span data-ttu-id="68fa4-256">Ce point d’extension ajoute une activation de complément contextuel sur un type d’entité spécifié.</span><span class="sxs-lookup"><span data-stu-id="68fa4-256">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="68fa4-257">Pour l’élément [VersionOverrides](versionoverrides.md) le contenant, l’attribut `xsi:type` doit avoir la valeur `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="68fa4-257">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="68fa4-258">Ce type d’élément est uniquement pris en charge par Outlook sur le web dans Office 365.</span><span class="sxs-lookup"><span data-stu-id="68fa4-258">This element type is only supported by Outlook on the web in Office 365.</span></span>

|  <span data-ttu-id="68fa4-259">Élément</span><span class="sxs-lookup"><span data-stu-id="68fa4-259">Element</span></span> |  <span data-ttu-id="68fa4-260">Description</span><span class="sxs-lookup"><span data-stu-id="68fa4-260">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="68fa4-261">Label</span><span class="sxs-lookup"><span data-stu-id="68fa4-261">Label</span></span>](#label) |  <span data-ttu-id="68fa4-262">Spécifie l’étiquette pour le complément dans la fenêtre contextuelle.</span><span class="sxs-lookup"><span data-stu-id="68fa4-262">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="68fa4-263">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="68fa4-263">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="68fa4-264">Spécifie l’URL de la fenêtre contextuelle.</span><span class="sxs-lookup"><span data-stu-id="68fa4-264">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="68fa4-265">Règle</span><span class="sxs-lookup"><span data-stu-id="68fa4-265">Rule</span></span>](rule.md) |  <span data-ttu-id="68fa4-266">Spécifie la ou les règles qui déterminent lorsqu’un complément s’active.</span><span class="sxs-lookup"><span data-stu-id="68fa4-266">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="68fa4-267">Étiquette</span><span class="sxs-lookup"><span data-stu-id="68fa4-267">Label</span></span>

<span data-ttu-id="68fa4-p115">Obligatoire. Libellé du groupe. L’attribut  **resid** doit être défini sur la valeur de l’attribut **id** d’un élément **String** dans l’élément **ShortStrings** de l’élément [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="68fa4-p115">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="68fa4-271">Exigences relatives à la mise en surbrillance</span><span class="sxs-lookup"><span data-stu-id="68fa4-271">Highlight requirements</span></span>

<span data-ttu-id="68fa4-p116">Le seul moyen pour qu’un utilisateur puisse activer un complément contextuel consiste à interagir avec une entité en surbrillance. Les développeurs peuvent contrôler les entités qui sont mises en surbrillance à l’aide de l’attribut `Highlight` de l’élément `Rule` pour les types de règles `ItemHasKnownEntity` et `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="68fa4-p116">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="68fa4-p117">Toutefois, il existe certaines limitations à connaître. Ces limitations sont en place pour vous assurer qu’il y aura toujours une entité en surbrillance dans les messages ou rendez-vous applicables pour permettre à l’utilisateur d’activer le complément.</span><span class="sxs-lookup"><span data-stu-id="68fa4-p117">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="68fa4-276">Les types d’entité `EmailAddress` et `Url` ne peuvent pas être mis en surbrillance et par conséquent ne peuvent pas être utilisés pour activer un complément.</span><span class="sxs-lookup"><span data-stu-id="68fa4-276">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="68fa4-277">Si vous utilisez une seule règle, la valeur `Highlight` DOIT être définie sur `all`.</span><span class="sxs-lookup"><span data-stu-id="68fa4-277">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="68fa4-278">Si vous utilisez un type de règle `RuleCollection` avec `Mode="AND"` pour combiner plusieurs règles, au moins l’une des règles DOIT définir `Highlight` sur la valeur `all`.</span><span class="sxs-lookup"><span data-stu-id="68fa4-278">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="68fa4-279">Si vous utilisez un type de règle `RuleCollection` avec `Mode="OR"` pour combiner plusieurs règles, toutes les règles DOIVENT définir `Highlight` sur la valeur `all`.</span><span class="sxs-lookup"><span data-stu-id="68fa4-279">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="68fa4-280">Exemple d’événement DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="68fa4-280">DetectedEntity event example</span></span>

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```