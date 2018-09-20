# <a name="javascript-api-for-office"></a>Interface API JavaScript pour Office

L’API JavaScript pour Office vous permet de créer des applications web qui interagissent avec les modèles objet dans des applications hôtes Office. Votre application font référence à la bibliothèque d’office.js, qui est un chargeur de script. La bibliothèque office.js charge les modèles d’objet qui s’appliquent à l’application Office qui exécute le complément. Vous pouvez utiliser les modèles d’objet JavaScript suivants :

- **API courantes** - API qui ont été introduites dans **Office 2013**. Cela est chargé pour **toutes les applications hôtes Office qui** et connecte votre application avec l’application cliente Office. Le modèle objet contient les API qui sont spécifiques aux clients Office et qui s’appliquent à plusieurs applications hôtes de client Office. Ce contenu est en cours **d’API partagé**. 

  **Outlook** utilise également la syntaxe de l’API courantes. Tout sous l’alias Office contient des objets que vous pouvez utiliser pour écrire des scripts qui interagissent avec le contenu de documents, feuilles de calcul, présentations, éléments de courrier et projets à partir de vos compléments Office. Vous devez utiliser ces API courantes si votre complément ciblera Office 2013 et versions ultérieures. Ce modèle d’objet utilise les rappels.

- **API spécifique à l’hôte** - API qui ont été introduites avec **Office 2016**. Ce modèle fournit des objets fortement typé spécifique à l’hôte qui correspondent aux objets que vous voyez lorsque vous utilisez des clients Office et représente l’avenir du JavaScript APIs Office familiers. Les API spécifiques à l’hôte sont actuellement l’API JavaScript de Word et l’interface API JavaScript Excel.

## <a name="supported-host-applications"></a>Applications hôtes prises en charge

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [Shared API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint et Project](requirement-sets/powerpoint-and-project-note.md) prennent en charge des compléments avec l’API JavaScript. Toutefois, elles actuellement n’ont pas API spécifique à l’hôte. Vous interagissez avec ces hôtes par le biais de l’API partagé.

En savoir plus sur les [hôtes pris en charge et les autres exigences](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins)

## <a name="open-api-specifications"></a>Spécifications d’ouverture de l’API

Au fur et à mesure que nous concevons et développons de nouvelles API pour les compléments Office, nous les mettons à votre disposition sur notre page de [spécifications d’ouverture de l’API](openspec.md) pour que vous puissiez fournir vos commentaires. Découvrez les nouvelles fonctionnalités dans le pipeline et donnez votre avis sur nos spécifications de conception.

## <a name="see-also"></a>Voir aussi

- [Référence d’API JavaScript Office](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)