# <a name="how-the-office-javascript-api-documentation-is-generated"></a>Comment la documentation de l’API JavaScript pour Office est générée

Les pages de documentation de référence JavaScript Office sont générées à partir de fichiers de définition de type et d’exemples d’extraits de code. Ce processus utilise un mélange d’outils open source et de scripts spécifiques au référentiel. Ce document vise à rendre les processus de ce référentiel transparents afin que la communauté puisse mieux tirer parti de ce contenu et y contribuer.

## <a name="content-sources"></a>Sources de contenu

Deux types de contenu sont combinés pour créer la documentation de référence Office-JS : les définitions de types et les extraits de code. Ceux-ci garantissent une couverture complète de l’API et donnent de petits exemples de code en ligne.

### <a name="type-definition-files"></a>Fichiers de définition de type

Les fichiers de définition de type sur [Definitely Typed](https://github.com/DefinitelyTyped/DefinitelyTyped) sont la seule source de vrai pour la documentation. Tout add-in Office qui utilise des compilations TypeScript à l’aide de ces fichiers de définition de type. Ces fonctionnalités offrent également aux développeurs JavaScript et TypeScript IntelliSense fonctionnalités. En construisant la documentation de référence à partir de ces définitions, nous fournissons des informations plus précises.

Il existe quatre fichiers d.ts pertinents qui fournissent du contenu source pour différentes sous-sections des documents.

- [office-js/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (Définitions de publication.)
  - [Excel (version)](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word (publication)](https://docs.microsoft.com/javascript/api/word_release)
  - [Sous-section OfficeExtensions de l’API commune](https://docs.microsoft.com/javascript/api/office)
- [office-js-preview/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (Définitions de l’aperçu.)
  - [Excel (prévisualisation)](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook (aperçu)](https://docs.microsoft.com/javascript/api/outlook)
  - [Word (prévisualisation)](https://docs.microsoft.com/javascript/api/word)
  - [API communes](https://docs.microsoft.com/javascript/api/office)
- [custom-functions-runtime/index.d.ts](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts) (définitions d’runtime des fonctions personnalisées Excel.)
  - [Fonctions personnalisées](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [office-runtime/index.d.ts](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts) (définitions d’office runtime pour la plateforme fonctions personnalisées.)
  - [Office Runtime](https://docs.microsoft.com/javascript/api/office-runtime)

Les versions antérieures des API ont leurs propres fichiers d.ts. Ceux-ci sont conservés lorsqu’un nouvel ensemble de conditions requises d’API est publié. Ils peuvent également être générés à l’aide de [l’outil De suppression de version.](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts) Ces anciens fichiers d.ts sont conservés de sorte que, dans le cas où les API sont corrigés ou modifiés, le comportement d’origine est toujours documenté. Cela est utile si vous devez cibler une version antérieure de l’API.

#### <a name="testing-type-definition-file-changes"></a>Test des modifications de fichier de définition de type

Les modifications apportées à la documentation pour l’API JavaScript office sont apportées en éditant les quatre fichiers d.ts mentionnés ci-dessus. Toutefois, vous pouvez tester une modification avant d’envoyer un pr à DefinitelyTyped (si vous devez, par exemple, tester comment votre mise en forme sera traduite en markdown) en éditant le fichier correspondant dans [generate-docs/script-inputs](https://github.com/OfficeDev/office-js-docs-reference/tree/master/generate-docs/script-inputs) et en exécutant [GenerateDocs.cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd). Lorsque vous y invitez, sélectionnez l’option « Fichiers locaux ».

Le fait d’apporter des modifications à une branche distante de ce repo entraîne la docs.microsoft.com de test à créer une branche de test. Cette branche est restituer sur review.docs.microsoft.com, qui est uniquement accessible par le personnel microsoft interne. Toute personne vérifiant votre pr vérifiera le site d’examen pour plus de précision.

### <a name="code-snippets"></a>Extraits de code

Des extraits de code sont ajoutés aux pages de référence à partir de deux sources :

- [Exemples de script lab](https://github.com/OfficeDev/office-js-snippets)
- [Extraits de code local](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

Les extraits de code locaux sont dans des fichiers yaml propres à l’hôte. Leur contenu est organisé par classe et par champ, afin qu’il puisse être mappé à l’endroit approprié dans une page de référence. Le langage de l’extrait de code (JavaScript ou TypeScript) est déduit par l’utilisation d’instructions await.

Les extraits script Lab sont extraits des exemples de travail. Actuellement, les exemples Excel, Outlook, PowerPoint et Word sont mappés pour référencer des sections de document par le biais de [fichiers de mappage.](https://github.com/OfficeDev/office-js-snippets/tree/prod/snippet-extractor-metadata) Ces méthodes correspondent à des exemples de méthodes individuels à des propriétés ou des méthodes dans l’API. Lorsque le référentiel office-js-snippets s’exécute, un fichier yaml contenant tous les extraits de code `yarn start` mappés est créé. [](https://github.com/OfficeDev/office-js-snippets/blob/prod/snippet-extractor-output/snippets.yaml) Ce fichier yaml est l’entrée dans les outils de documentation de référence.

## <a name="tooling-pipeline"></a>Pipeline d’outils

![Image montrant le flux de contrôle entre le préprocesseur, l’extracteur d’API, le midprocessor, l’document d’API et le postprocesseur.](ToolingPipeline.png)

Entre les sources de contenu et les pages finales, le contenu de la documentation passe par cinq étapes d’outils :

1. [Script de préprocesseur](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [Extracteur d’API](https://api-extractor.com/)
1. [Script Midprocessor](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [Document d’API](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [Script postprocesseur](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

Le préprocesseur prend les fichiers d.ts et les divise en sections propres à l’hôte. Il effectue tout nettoyage nécessaire aux outils suivants pour traiter correctement les données.

L’extracteur d’API convertit les fichiers d.ts en données JSON. Cela permet de tokenize toutes les données de type, ce qui facilite l’analyse.

Le midprocessor récupère les extraits de code et les associe aux hôtes appropriés, et nettoie les liens entre les objets Outlook et les objets d’API communes.

L’document d’API convertit les données JSON en fichiers .yml. Les fichiers .yml sont convertis en markdown par le système de publication Open qui publie nos documents dans docs.microsoft.com. L’document d’API contient également une extension spécifique d’Office qui insère nos extraits de code.

Le postprocesseur nettoie la table des matières et déplace les fichiers .yml dans le [dossier de publication.](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen)

Ces cinq étapes sont effectuées lors de l’utilisation [de GenerateDocs.cmd.](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) Ce script gère également l’installation du module de nœud, nettoie les anciens ensembles de fichiers et les fichiers de définition de type des versions pour chaque ensemble de conditions requises.
