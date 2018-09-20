# <a name="highresolutioniconurl-element"></a>HighResolutionIconUrl, élément

Spécifie l’URL de l’image qui est utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store sur les écrans à haute résolution (DPI).

**Type de complément :** Application de contenu, de volet Office, de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>Peut contenir

[Override](override.md)

## <a name="attributes"></a>Attributs

|**Attribut**|**Type**|**Obligatoire**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|chaîne (URL)|obligatoire|Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).|

## <a name="remarks"></a>Remarques

Pour un complément de messagerie, l’icône apparaît dans l’interface utilisateur, sous **Fichier**  >  **Gérer les compléments**. Pour un complément de contenu ou de volet Office, l’icône apparaît dans l’interface utilisateur, sous **Insérer**  >  **Compléments**.

L’image doit être dans l’un des formats de fichier suivants, avec une résolution recommandée de 64 x 64 pixels : GIF, JPG, PNG, EXIF, BMP ou TIFF. Pour plus d’informations, voir la section _créer une identité visuelle cohérente pour votre application_ de [créer des annonces efficaces dans AppSource et dans Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).
