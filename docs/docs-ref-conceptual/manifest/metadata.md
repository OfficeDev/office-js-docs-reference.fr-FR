# <a name="metadata-element"></a>Metadata, élément

Définit les paramètres de métadonnées utilisés par une fonction personnalisée dans Excel.

## <a name="attributes"></a>Attributs

Aucune

## <a name="child-elements"></a>Éléments enfants

|  Élément  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Oui  | Chaîne de l’id de ressource du fichier JSON utilisé par des fonctions personnalisées. |

## <a name="example"></a>Exemples

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
