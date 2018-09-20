# <a name="set-element"></a>Élément Set

Spécifie un ensemble de conditions requises de l’API JavaScript pour Office nécessaires à l’activation de votre complément Office.

**Type de complément :** Application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>Contenu dans

[Ensembles](sets.md)

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|Nom|string|obligatoire|Nom d’un [ensemble de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).|
|MinVersion|chaîne|facultatif|Spécifie la version minimale de l’ensemble d’API requis par votre complément. Remplace la valeur de **DefaultMinVersion**, si elle est spécifiée dans l’élément parent [Sets](sets.md).|

## <a name="remarks"></a>Remarques

Pour plus d’informations sur les ensembles de ressources, voir [définit les versions d’Office et de la spécification](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Pour plus d’informations sur l’attribut **MinVersion** de l’élément **Set** et sur l’attribut **DefaultMinVersion** de l’élément **Sets**, voir l’article relatif à la [définition de l’élément Requirements dans le manifeste](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

> [!IMPORTANT] 
> Pour les compléments de messagerie, il n’existe qu’une `"Mailbox"` exigence ensemble disponible. Cet ensemble de conditions requises contient le sous-ensemble ensemble d’API pris en charge dans les compléments de messagerie pour Outlook, et vous devez spécifier le `"Mailbox"` exigence défini dans le manifeste de votre complément messagerie (il n’est pas facultatif comme c’est le cas pour les tâches et de contenu compléments volet). En outre, vous ne pouvez pas déclarer prise en charge des méthodes spécifiques dans des compléments de messagerie.
