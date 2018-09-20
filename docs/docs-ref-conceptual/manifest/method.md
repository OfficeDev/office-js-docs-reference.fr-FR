# <a name="method-element"></a>Élément Method

Spécifie une méthode individuelle de l’API JavaScript pour Office requise pour l’activation de votre complément Office.

**Type de complément :** Application de contenu et de volet Office

## <a name="syntax"></a>Syntaxe

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Contenu dans

[Méthodes](methods.md)

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|Nom|string|obligatoire|Spécifie le nom de la méthode qualifiée requise avec son objet parent. Par exemple, pour spécifier la méthode **getSelectedDataAsync**, vous devez spécifier `"Document.getSelectedDataAsync"`.|

## <a name="remarks"></a>Remarques

Les éléments de **méthodes** et de la **méthode** ne sont pas pris en charge par les compléments de messagerie. Pour plus d’informations sur les ensembles de ressources, voir [définit les versions d’Office et de la spécification](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

> [!IMPORTANT] 
> Car il n’existe aucun moyen pour spécifier la version minimale requise pour les différentes méthodes, pour vous assurer qu’une méthode est disponible à l’exécution, vous devez également utiliser une instruction **if** lors de l’appel de cette méthode dans le script de votre complément. Pour plus d’informations, voir [Présentation de l’API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

