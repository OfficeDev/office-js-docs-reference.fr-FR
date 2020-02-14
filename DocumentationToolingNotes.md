# <a name="how-the-office-javascript-api-documentation-is-generated"></a>Comment la documentation de l’API JavaScript pour Office est générée

Les pages de documentation de référence JavaScript Office sont générées à partir de fichiers de définition de type et d’extraits d’exemples. Ce processus utilise un mélange d’outils open source et de scripts spécifiques au référentiel. Ce document a pour but de simplifier les processus de ce référentiel afin que la Communauté puisse tirer le meilleur parti et contribuer à ce contenu.

## <a name="content-sources"></a>Sources de contenu

Deux types de contenu sont combinés pour créer la documentation de référence Office-JS : définitions de types et extraits de code. Elles garantissent une couverture complète de l’API et donnent des exemples de code incorporés de petite taille.

### <a name="type-definition-files"></a>Fichiers de définition de type

Les fichiers de définition de type sur la base d’une [saisie définitive](https://github.com/DefinitelyTyped/DefinitelyTyped) représentent la seule source de vérité pour la documentation. N’importe quel complément Office qui utilise une compilation d’appel à l’aide de ces fichiers de définition de type. Elles donnent également aux développeurs JavaScript et de la machine à écrire des fonctionnalités IntelliSense. En construisant la documentation de référence à partir de ces définitions, nous fournissons des informations plus précises.

Il existe quatre fichiers d. TS pertinents qui fournissent le contenu source de différentes sous-sections des documents.

- [Office-js/index. d. TS](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (définitions de publication)
  - [Excel (version Release)](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word (publication)](https://docs.microsoft.com/javascript/api/word_release)
  - [Sous-section OfficeExtensions de l’API commune](https://docs.microsoft.com/javascript/api/office)
- [Office-js-aperçu/index. d. TS](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (les définitions d’aperçu.)
  - [Excel (aperçu)](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook (aperçu)](https://docs.microsoft.com/javascript/api/outlook)
  - [Word (aperçu)](https://docs.microsoft.com/javascript/api/word)
  - [API communes](https://docs.microsoft.com/javascript/api/office)
- [Custom-Functions-Runtime/index. d. TS](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts) (les définitions d’exécution des fonctions personnalisées Excel.)
  - [Fonctions personnalisées](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [Office-Runtime/index. d. TS](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts) (les définitions d’exécution Office pour la plateforme de fonctions personnalisées.)
  - [Office Runtime](https://docs.microsoft.com/javascript/api/office-runtime)

Les versions antérieures des API ont leurs propres fichiers d. TS. Ces éléments sont conservés lorsqu’un nouvel ensemble de conditions requises de l’API est publié. Ils peuvent également être générés à l’aide de l' [outil de suppression de version](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts). Ces anciens fichiers d. TS sont conservés de sorte que, dans les API d’événements, les correctifs ou les modifications soient toujours traités. Cette fonctionnalité est utile si vous devez cibler une version plus ancienne de l’API.

#### <a name="testing-type-definition-file-changes"></a>Test des modifications apportées au fichier de définition de type

Toutes les modifications apportées à la documentation pour l’API JavaScript Office sont effectuées en modifiant les quatre fichiers de la version d. TS mentionnés ci-dessus. Toutefois, vous pouvez tester une modification avant d’envoyer un PR à DefinitelyTyped (si vous le souhaitez, par exemple, tester le mode de conversion de votre mise en forme en démarque) en modifiant le fichier correspondant dans [generate-docs/script-inputs](https://github.com/OfficeDev/office-js-docs-reference/tree/master/generate-docs/script-inputs) et en exécutant [GenerateDocs. cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd). Lorsque vous y êtes invité, sélectionnez l’option « fichiers locaux ».

Si vous poussez les modifications vers une branche distante de cette référentiel, la plateforme docs.microsoft.com crée une branche de test. Cette branche est rendue sur review.docs.microsoft.com, qui est uniquement accessible par le personnel de Microsoft interne. Toute personne qui examine votre PR vérifiera la précision du site de révision.

### <a name="code-snippets"></a>Extraits de code

Des extraits de code d’exemple de code sont ajoutés aux pages de référence à partir de deux sources :

- [Exemples de laboratoire de script](https://github.com/OfficeDev/office-js-snippets)
- [Extraits de code locaux](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

Les extraits de code locaux se trouvent dans des fichiers YAML spécifiques à l’hôte. Leur contenu est organisé par classe et par champ, de sorte qu’il peut être mappé à l’emplacement approprié dans une page de référence. La langue de l’extrait de code (JavaScript ou machine à écrire) est déduite par l’utilisation d’instructions await.

Les extraits de code d’atelier de script sont extraits des exemples de travail. Actuellement, les exemples Excel et Word sont mappés aux sections de document de référence par le biais [d’une paire de fichiers de mappage](https://github.com/OfficeDev/office-js-snippets/tree/master/snippet-extractor-metadata). Ces méthodes correspondent aux propriétés ou méthodes de l’API. Lorsque le référentiel Office-js-snippets est `yarn start` exécuté, [un fichier YAML](https://github.com/OfficeDev/office-js-snippets/blob/master/snippet-extractor-output/snippets.yaml) contenant tous les extraits de code mappés est créé. Ce fichier YAML est l’entrée dans les outils de documentation de référence.

## <a name="tooling-pipeline"></a>Pipeline d’outils

![Image illustrant le flux de contrôle à partir de la saisie définitive, au préprocesseur, à l’extracteur d’API, à midprocessor, au documenteur de l’API et par rapport au postprocesseur.](ToolingPipeline.png)

Entre les sources de contenu et les pages finales, le contenu de la documentation passe par cinq étapes d’outils :

1. [Script du préprocesseur](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [Extracteur d’API](https://api-extractor.com/)
1. [Script Midprocessor](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [Documentation de l’API](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [Script de postprocesseur](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

Le préprocesseur prend les fichiers d. TS et les divise en sections spécifiques à l’hôte. Il effectue tout nettoyage nécessaire aux outils suivants pour traiter correctement les données.

L’extracteur d’API convertit les fichiers d. TS en données JSON. Cela régit toutes les données de type, ce qui facilite l’analyse.

Le midprocessor récupère les extraits de code et les associe aux hôtes appropriés.

L’API documenter convertit les données JSON en fichiers. yml. Les fichiers. yml sont convertis en démarques par le système de publication ouvert qui publie nos documents sur docs.microsoft.com. L’API documenter contient également une extension spécifique à Office qui insère nos extraits de code.

Le postprocesseur nettoie la table des matières et déplace les fichiers. yml dans le dossier de [publication](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen).

Ces cinq étapes sont exécutées lors de l’exécution de [GenerateDocs. cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) . Ce script gère également l’installation du module de nœud, nettoie les anciens ensembles de fichiers et les fichiers de définition de type versions pour chaque ensemble de conditions requises.
