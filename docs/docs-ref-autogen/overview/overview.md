---
layout: LandingPage
ms.topic: landing-page
title: Référence de l’API JavaScript d’Office
description: Les API Office JavaScript par hôte et version.
author: o365devx
ms.author: o365devx
ms.prod: non-product-specific
localization_priority: Priority
ms.date: 06/17/2020
ms.openlocfilehash: aff744f62d55449200a821634510ac3da5ea41c0
ms.sourcegitcommit: e94c95582f58781bf193461f1b8148fac833dba0
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/22/2020
ms.locfileid: "45365089"
---
# <a name="office-add-ins-javascript-api-reference"></a>Référence de l’API JavaScript pour les compléments Office

L’interface API JavaScript pour Office vous permet de créer des applications web qui interagissent avec les modèles objet dans les applications hôtes Office. Utilisez cette section pour en savoir plus sur les classes, les méthodes et les autres types disponibles pour créer des compléments Office.

La liste suivante répertorie les API pour les [applications hôtes Office prises en charge](/office/dev/add-ins/overview/office-add-in-availability). Les liens API courants incluent toutes les API non attribuées à un hôte particulier (comme indiqué dans [Configuration requise pour l’API commune d’Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)). D'autres éléments renvoient à une version de la documentation de référence de l'API pour cet hôte en fonction d’un ensemble de conditions requises. La documentation de référence est versionnée pour inclure toutes les API jusqu'à cet ensemble de conditions requises (par exemple, ExcelApi 1.3 montre les API dans ExcelApi 1.1, 1.2, 1.3 ainsi que les API courantes).

`ExcelApiOnline 1.1` est un ensemble de conditions requises particulier. Il contient les dernières API pour Excel sur le Web, mais ces API ne sont peut-être pas encore entièrement prises en charge sur toutes les plateformes. Pour plus d’informations, consultez [Ensemble de conditions requises pour l’API JavaScript Excel en ligne uniquement](/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set).

> [!TIP]
> Vous pouvez modifier la version d’une page de référence à l’aide du menu déroulant sélection de filtre au-dessus de la table des matières à tout moment. Si la page n’existe pas dans cette version, vous revenez à la version actuelle.

<h2>Hôtes Office</h2>

<ul class="cardsK panelContent cols cols3">
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-excel.svg" alt="Excel add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>API Excel</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-preview">Aperçu ExcelApi</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-online">ExcelApiOnline 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.11">ExcelApi 1.11</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.10">ExcelApi 1.10</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.9">ExcelApi 1.9</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.8">ExcelApi 1.8</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.7">ExcelApi 1.7</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.6">ExcelApi 1.6</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.5">ExcelApi 1.5</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.4">ExcelApi 1.4</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.3">ExcelApi 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.2">ExcelApi 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/excel?view=excel-js-1.1">ExcelApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=excel-js-preview">API communes</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-outlook.svg" alt="Outlook add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>APIs Outlook</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-preview">Aperçu de la boîte aux lettres</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.8">Mailbox 1.8</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.7">Mailbox 1.7</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.6">Mailbox 1.6</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.5">Mailbox 1.5</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.4">Mailbox 1.4</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.3">Mailbox 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.2">Mailbox 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/outlook?view=outlook-js-1.1">Mailbox 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=outlook-js-preview">API communes</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-word.svg" alt="Word add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>API Word</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-preview">Aperçu WordApi</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.3">WordApi 1.3</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.2">WordApi 1.2</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/word?view=word-js-1.1">WordApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=word-js-preview">API communes</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-onenote.svg" alt="OneNote add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>API OneNote</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/onenote?view=onenote-js-1.1">OneNoteApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=onenote-js-1.1">API communes</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-visio.svg" alt="Visio add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>API Visio</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/visio?view=visio-js-1.1">VisioApi 1.1</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-powerpoint.svg" alt="PowerPoint add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>API PowerPoint</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/powerpoint?view=powerpoint-js-1.1">PowerPointApi 1.1</a></li>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=powerpoint-js-1.1">API communes</a></li>
                </ul>
            </div>
        </a>
    </li>
    <li>
        <a class="card x-hidden-focus">
            <div class="cardImageOuter">
                <div class="cardImage">
                    <img src="/javascript/api/overview/images/logo-project.svg" alt="Project add-ins" />
                </div>
            </div>
            <div class="cardText">
                <h3>API Project</h3>
                <ul>
                    <li><a style="font-size: 1rem;" href="/javascript/api/office?view=common-js">API communes uniquement</a></li>
                </ul>
            </div>
        </a>
    </li>
</ul>

> [!NOTE]
> Si vous recherchez les API JavaScript pour le développement de scripts Office, visitez la page [Référence de l'API Office Scripts](/javascript/api/office-scripts/overview).

## <a name="see-also"></a>Voir aussi

- [À propos des compléments Office](/office/dev/add-ins/overview)
- [Disponibilité des compléments Office sur les plateformes et les hôtes](/office/dev/add-ins/overview/office-add-in-availability)
- [Versions d’Office et ensembles de conditions requises](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Explorer l’API JavaScript Office à l’aide de Script Lab](/office/dev/add-ins/overview/explore-with-script-lab)
