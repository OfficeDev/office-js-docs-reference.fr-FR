# <a name="desktopformfactor-element"></a><span data-ttu-id="2aca4-101">Élément DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="2aca4-101">DesktopFormFactor element</span></span>

<span data-ttu-id="2aca4-p101">Spécifie les paramètres d’un complément en fonction du facteur de forme pour bureau. Le facteur de forme pour bureau inclut Office pour Windows, Office pour Mac et Office Online. Il contient toutes les informations de complément pour ce facteur de forme à l’exception du nœud **Resources**.</span><span class="sxs-lookup"><span data-stu-id="2aca4-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="2aca4-p102">Chaque définition de facteur de forme pour bureau contient l’élément **FunctionFile** et au moins un élément **ExtensionPoint**. Pour plus d’informations, voir [Élément FunctionFile](functionfile.md) et [Élément ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="2aca4-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span> 

## <a name="child-elements"></a><span data-ttu-id="2aca4-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="2aca4-107">Child elements</span></span>

| <span data-ttu-id="2aca4-108">Élément</span><span class="sxs-lookup"><span data-stu-id="2aca4-108">Element</span></span>                               | <span data-ttu-id="2aca4-109">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="2aca4-109">Required</span></span> | <span data-ttu-id="2aca4-110">Description</span><span class="sxs-lookup"><span data-stu-id="2aca4-110">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="2aca4-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="2aca4-111">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="2aca4-112">Oui</span><span class="sxs-lookup"><span data-stu-id="2aca4-112">Yes</span></span>      | <span data-ttu-id="2aca4-113">Définit l’emplacement où se trouvent les fonctionnalités d’un complément</span><span class="sxs-lookup"><span data-stu-id="2aca4-113">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="2aca4-114">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="2aca4-114">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="2aca4-115">Oui</span><span class="sxs-lookup"><span data-stu-id="2aca4-115">Yes</span></span>      | <span data-ttu-id="2aca4-116">URL pointant vers un fichier qui contient les fonctions JavaScript.</span><span class="sxs-lookup"><span data-stu-id="2aca4-116">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="2aca4-117">GetStarted</span><span class="sxs-lookup"><span data-stu-id="2aca4-117">GetStarted</span></span>](getstarted.md)         | <span data-ttu-id="2aca4-118">Non</span><span class="sxs-lookup"><span data-stu-id="2aca4-118">No</span></span>       | <span data-ttu-id="2aca4-119">Définit la légende qui s’affiche lorsque vous installez le complément dans des hôtes Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="2aca4-119">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="2aca4-120">Exemple pour DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="2aca4-120">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
