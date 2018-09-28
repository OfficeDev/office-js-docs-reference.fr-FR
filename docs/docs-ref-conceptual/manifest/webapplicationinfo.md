# <a name="webapplicationinfo-element"></a><span data-ttu-id="1fd72-101">Élément WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="1fd72-101">WebApplicationInfo element</span></span>

<span data-ttu-id="1fd72-102">Prend en charge l’authentification unique (SSO) dans des compléments Office. Cet élément contient des informations sur le complément sous deux formes :</span><span class="sxs-lookup"><span data-stu-id="1fd72-102">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="1fd72-103">En tant que *ressource* OAuth 2.0 pour laquelle l’application Office peut requérir des autorisations.</span><span class="sxs-lookup"><span data-stu-id="1fd72-103">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="1fd72-104">Un *client* OAuth 2.0 pouvant requérir des autorisations dans Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="1fd72-104">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

<span data-ttu-id="1fd72-105">**WebApplicationInfo** est un élément enfant de l’élément [VersionOverrides](versionoverrides.md) dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="1fd72-105">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="1fd72-106">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="1fd72-106">Child elements</span></span>

|  <span data-ttu-id="1fd72-107">Élément</span><span class="sxs-lookup"><span data-stu-id="1fd72-107">Element</span></span> |  <span data-ttu-id="1fd72-108">Obligatoire</span><span class="sxs-lookup"><span data-stu-id="1fd72-108">Required</span></span>  |  <span data-ttu-id="1fd72-109">Description</span><span class="sxs-lookup"><span data-stu-id="1fd72-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1fd72-110">**Id**</span><span class="sxs-lookup"><span data-stu-id="1fd72-110">**Id**</span></span>    |  <span data-ttu-id="1fd72-111">Oui</span><span class="sxs-lookup"><span data-stu-id="1fd72-111">Yes</span></span>   |  <span data-ttu-id="1fd72-112">**ID d’application** du service associé au complément, tel qu’inscrit dans le point de terminaison Azure Active Directory (Azure AD) v2.0.</span><span class="sxs-lookup"><span data-stu-id="1fd72-112">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="1fd72-113">**Resource**</span><span class="sxs-lookup"><span data-stu-id="1fd72-113">**Resource**</span></span>  |  <span data-ttu-id="1fd72-114">Oui</span><span class="sxs-lookup"><span data-stu-id="1fd72-114">Yes</span></span>   |  <span data-ttu-id="1fd72-115">Spécifie l’**URI de l’ID d’application** du complément, tel qu’inscrit dans le point de terminaison Azure AD v2.0.</span><span class="sxs-lookup"><span data-stu-id="1fd72-115">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="1fd72-116">Scopes</span><span class="sxs-lookup"><span data-stu-id="1fd72-116">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="1fd72-117">Non</span><span class="sxs-lookup"><span data-stu-id="1fd72-117">No</span></span>  |  <span data-ttu-id="1fd72-118">Spécifie les autorisations dont le complément a besoin pour Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="1fd72-118">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="1fd72-119">Actuellement, il est nécessaire que les ressources de votre complément correspond à son hôte.</span><span class="sxs-lookup"><span data-stu-id="1fd72-119">Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="1fd72-120">Office ne demandera pas un jeton pour un complément à moins de pouvoir prouver qu’il en est le propriétaire ; à l’heure actuelle, ceci s’effectue en hébergeant le complément sous le nom de domaine complet de la ressource.</span><span class="sxs-lookup"><span data-stu-id="1fd72-120">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="1fd72-121">Exemple pour WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="1fd72-121">WebApplicationInfo example</span></span>

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
