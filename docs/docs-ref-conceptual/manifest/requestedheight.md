# <a name="requestedheight-element"></a>Élément RequestedHeight.

Spécifie la hauteur initiale (en pixels) d’un complément de contenu ou le complément messagerie. 

**Type complément :** Contenu de messagerie

## <a name="syntax"></a>Syntaxe

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>Contenu dans

- [DefaultSettings](defaultsettings.md) (Contenu compléments) avec une valeur qui peut être comprise entre 32 et 1 000
- [DesktopSettings](desktopsettings.md) et [TabletSettings](tabletsettings.md) (compléments messagerie) avec une valeur qui peut être comprise entre 32 et 450
- [ExtensionPoint](extensionpoint.md) (Messagerie contextuelle des compléments) avec une valeur qui peut être entre 140 et 450 pour le point d’extension **DetectedEntity** et entre 32 et 450 pour le point d’extension **CustomPane**