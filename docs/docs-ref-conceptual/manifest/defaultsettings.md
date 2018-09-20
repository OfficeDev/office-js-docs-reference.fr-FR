# <a name="defaultsettings-element"></a>Élément DefaultSettings

Spécifie l’emplacement de la source par défaut et d’autres paramètres par défaut pour le contenu ou le complément de volet de tâches.

**Type de complément :** Application de contenu et de volet Office

## <a name="syntax"></a>Syntaxe

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>Contenu dans

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Peut contenir

|**Élément**|**Content**|**Courrier**|**Volet Office**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Remarques

L’emplacement source et les autres paramètres de l’élément **DefaultSettings** s’appliquent uniquement aux compléments de volet Office et de contenu. Pour les compléments de messagerie, vous spécifiez les emplacements par défaut pour les fichiers sources et d’autres paramètres par défaut dans l’élément [FormSettings](formsettings.md).

