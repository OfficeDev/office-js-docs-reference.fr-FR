# <a name="namespace-element"></a>Élément Namespace

Définit l’espace de noms utilisé par une fonction personnalisée dans Excel.

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **RESID = « espace de noms »**  |  Oui  | Doit correspondre à titre ShortStrings votre fonction personnalisée, spécifié dans l’élément de [ressources](resources.md) . |

## <a name="child-elements"></a>Éléments enfants

Aucune

## <a name="example"></a>Exemples

```xml
<Namespace resid="namespace" />
```
