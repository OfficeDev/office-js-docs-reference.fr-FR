# <a name="sourcelocation-element"></a>Élément SourceLocation

Spécifie les emplacements des fichiers source pour votre complément Office sous forme d’URL comprenant entre 1 et 2 018 caractères. L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.

**Type de complément :** Application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>Contenu dans

- [DefaultSettings](defaultsettings.md) (compléments de contenu et de volet Office)
- [FormSettings](formsettings.md) (compléments de messagerie)
- [ExtensionPoint](extensionpoint.md) (compléments de messagerie contextuels)

## <a name="can-contain"></a>Peut contenir

[Override](override.md)

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|obligatoire|Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).|
