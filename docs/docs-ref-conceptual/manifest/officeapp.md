# <a name="officeapp-element"></a>OfficeApp, élément

Élément racine dans le manifeste d’un complément Office.

**Type de complément :** Application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a>Contenu dans

 _none_

## <a name="must-contain"></a>Doit contenir

|**Élément**|**Content**|**Courrier**|**Volet Office**|
|:-----|:-----|:-----|:-----|
|
  [Id](id.md)|x|x|x|
|[Version](version.md)|x|x|x|
|[ProviderName](providername.md)|x|x|x|
|[DefaultLocale](defaultlocale.md)|x|x|x|
|[DefaultSettings](defaultsettings.md)|x||x|
|[DisplayName](displayname.md)|x|x|x|
|[Description](description.md)|x|x|x|
|[FormSettings](formsettings.md)||x||
|[Autorisations](permissions.md)|x||x|
|[Règle](rule.md)||x||

## <a name="can-contain"></a>Peut contenir

|**Élément**|**Content**|**Courrier**|**Volet Office**|
|:-----|:-----|:-----|:-----|
|[AlternateId](alternateid.md)|x|x|x|
|[IconUrl](iconurl.md)|x|x|x|
|[HighResolutionIconUrl](highresolutioniconurl.md)|x|x|x|
|[SupportUrl](supporturl.md)|x|x|x|
|[AppDomains](appdomains.md)|x|x|x|
|[Hôtes](hosts.md)|x|x|x|
|[Configuration requise](requirements.md)|x|x|x|
|[AllowSnapshot](allowsnapshot.md)|x|||
|[Autorisations](permissions.md)||x||
|[DisableEntityHighlighting](disableentityhighlighting.md)||x||
|[Dictionnaire](dictionary.md)|||x|
|[VersionOverrides](versionoverrides.md)|X |X |X |

## <a name="attributes"></a>Attributs

|||
|:-----|:-----|
|xmlns|Définit la version de schéma et l’espace de noms du manifeste de complément Office. Cet attribut doit toujours être défini sur `"http://schemas.microsoft.com/office/appforoffice/1.1"`.|
|xmlns:xsi|Définit l’instance XMLSchema. Cet attribut doit toujours être défini sur `"http://www.w3.org/2001/XMLSchema-instance"`.|
|xsi:type|Définit le type de complément Office. Cet attribut doit être défini sur l’une des options suivantes : `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`.|
