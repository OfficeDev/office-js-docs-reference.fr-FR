# <a name="allformfactors-element"></a>Élément AllFormFactors

Spécifie les paramètres d’un complément pour tous les facteurs de forme. Actuellement, la seule fonctionnalité à l’aide de **AllFormFactors** est fonctions personnalisées. **AllFormFactors** est un élément requis lors de l’utilisation des fonctions personnalisées.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  Oui |  Définit l’emplacement où se trouvent les fonctionnalités d’un complément |

## <a name="allformfactors-example"></a>Exemple AllFormFactors

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
