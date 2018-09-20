# <a name="requirements-element"></a>Élément Requirements

Spécifie l’ensemble minimal des conditions requises de l’API JavaScript pour Office ([ensembles des conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) et/ou méthodes) que votre complément Office doit activer.

**Type de complément :** Application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Peut contenir

|**Élément**|**Content**|**Courrier**|**Volet Office**|
|:-----|:-----|:-----|:-----|
|[Ensembles](sets.md)|x|x|x|
|[Méthodes](methods.md)|x||x|

## <a name="remarks"></a>Remarques

Pour plus d’informations sur les ensembles de ressources, voir [définit les versions d’Office et de la spécification](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

