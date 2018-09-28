# <a name="appdomains-element"></a><span data-ttu-id="d4b1a-101">Élément AppDomains</span><span class="sxs-lookup"><span data-stu-id="d4b1a-101">AppDomains element</span></span>

<span data-ttu-id="d4b1a-p101">Répertorie tout domaine supplémentaire qui sera utilisé par votre complément Office pour charger des pages en plus du domaine spécifié dans l’élément SourceLocation. Pour chaque domaine supplémentaire, indiquez un élément AppDomain.</span><span class="sxs-lookup"><span data-stu-id="d4b1a-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="d4b1a-104">**Type de complément :** Application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="d4b1a-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d4b1a-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="d4b1a-105">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

## <a name="contained-in"></a><span data-ttu-id="d4b1a-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="d4b1a-106">Contained in</span></span>

[<span data-ttu-id="d4b1a-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="d4b1a-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="d4b1a-108">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="d4b1a-108">Can contain</span></span>

[<span data-ttu-id="d4b1a-109">AppDomain</span><span class="sxs-lookup"><span data-stu-id="d4b1a-109">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="d4b1a-110">Remarques</span><span class="sxs-lookup"><span data-stu-id="d4b1a-110">Remarks</span></span>

<span data-ttu-id="d4b1a-p102">Par défaut, votre complément peut charger n’importe quelle page qui se trouve dans le même domaine que l’emplacement indiqué dans l’élément **SourceLocation**. Pour charger des pages qui ne sont pas dans le même domaine que le complément, spécifiez les domaines à l’aide des éléments **AppDomains** et **AppDomain**. Vous devez indiquer une valeur pour cet élément.</span><span class="sxs-lookup"><span data-stu-id="d4b1a-p102">By default, your add-in can load any page that is in the same domain as the location specified in the **SourceLocation** element. To load pages that are not in the same domain as the add-in, specify the domains by using the **AppDomains** and **AppDomain** elements. This element can't be empty.</span></span> 
