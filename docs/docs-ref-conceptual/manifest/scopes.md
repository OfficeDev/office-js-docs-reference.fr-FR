# <a name="scopes-element"></a><span data-ttu-id="e5641-101">Élément Scope</span><span class="sxs-lookup"><span data-stu-id="e5641-101">Scopes element</span></span>

<span data-ttu-id="e5641-102">Contient des autorisations Microsoft Graph requises par le complément.</span><span class="sxs-lookup"><span data-stu-id="e5641-102">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="e5641-103">L’Office Store se sert de l’élément Scope pour créer une boîte de dialogue de consentement.</span><span class="sxs-lookup"><span data-stu-id="e5641-103">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="e5641-104">Lorsque les utilisateurs installent le complément à partir du Store, ils sont invités à lui accorder les autorisations spécifiées à leurs données Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="e5641-104">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="e5641-105">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="e5641-105">Child elements</span></span>

|  <span data-ttu-id="e5641-106">Élément</span><span class="sxs-lookup"><span data-stu-id="e5641-106">Element</span></span> |  <span data-ttu-id="e5641-107">Type</span><span class="sxs-lookup"><span data-stu-id="e5641-107">Type</span></span>  |  <span data-ttu-id="e5641-108">Description</span><span class="sxs-lookup"><span data-stu-id="e5641-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e5641-109">**Scope**</span><span class="sxs-lookup"><span data-stu-id="e5641-109">**Scope**</span></span>                |  <span data-ttu-id="e5641-110">string</span><span class="sxs-lookup"><span data-stu-id="e5641-110">string</span></span>     |   <span data-ttu-id="e5641-111">Nom d’une autorisation Microsoft Graph ; par exemple, Files.Read.All.</span><span class="sxs-lookup"><span data-stu-id="e5641-111">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="e5641-112">Exemples</span><span class="sxs-lookup"><span data-stu-id="e5641-112">Example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
