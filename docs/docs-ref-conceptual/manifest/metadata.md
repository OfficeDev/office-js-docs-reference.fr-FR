# <a name="metadata-element"></a><span data-ttu-id="25a11-101">Metadata, élément</span><span class="sxs-lookup"><span data-stu-id="25a11-101">Metadata element</span></span>

<span data-ttu-id="25a11-102">Définit les paramètres de métadonnées utilisés par une fonction personnalisée dans Excel.</span><span class="sxs-lookup"><span data-stu-id="25a11-102">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="25a11-103">Attributs</span><span class="sxs-lookup"><span data-stu-id="25a11-103">Attributes</span></span>

<span data-ttu-id="25a11-104">Aucune</span><span class="sxs-lookup"><span data-stu-id="25a11-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="25a11-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="25a11-105">Child elements</span></span>

|  <span data-ttu-id="25a11-106">Élément</span><span class="sxs-lookup"><span data-stu-id="25a11-106">Element</span></span>  |  <span data-ttu-id="25a11-107">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="25a11-107">Required</span></span>  |  <span data-ttu-id="25a11-108">Description</span><span class="sxs-lookup"><span data-stu-id="25a11-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="25a11-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="25a11-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="25a11-110">Oui</span><span class="sxs-lookup"><span data-stu-id="25a11-110">Yes</span></span>  | <span data-ttu-id="25a11-111">Chaîne de l’id de ressource du fichier JSON utilisé par des fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="25a11-111">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="25a11-112">Exemples</span><span class="sxs-lookup"><span data-stu-id="25a11-112">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
