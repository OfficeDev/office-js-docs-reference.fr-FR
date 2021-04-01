# <a name="office-javascript-api-reference"></a><span data-ttu-id="db43a-101">Référence de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="db43a-101">Office JavaScript API Reference</span></span>

<span data-ttu-id="db43a-102">Bienvenue dans le référentiel de documentation de référence de l’API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="db43a-102">Welcome to the Office JavaScript API Reference documentation repository.</span></span> <span data-ttu-id="db43a-103">Pour une expérience optimale, nous vous recommandons de consulter ce contenu sur [docs.microsoft.com](https://docs.microsoft.com/javascript/api/overview/office).</span><span class="sxs-lookup"><span data-stu-id="db43a-103">For the best experience, we recommend you view this content on [docs.microsoft.com](https://docs.microsoft.com/javascript/api/overview/office).</span></span>

> <span data-ttu-id="db43a-104">**Remarque**: vous trouverez les fichiers sources de documentation pour les concepts de l’API JavaScript pour Office, les démarrages rapides, les didacticiels et les guides pratiques dans le référentiel [OfficeDev/office-js-docs-pr.](https://github.com/OfficeDev/office-js-docs-pr)</span><span class="sxs-lookup"><span data-stu-id="db43a-104">**Note**: You can find the documentation source files for Office JavaScript API concepts, quick starts, tutorials, and how-to guides in the [OfficeDev/office-js-docs-pr](https://github.com/OfficeDev/office-js-docs-pr) repository.</span></span>

## <a name="give-us-your-feedback"></a><span data-ttu-id="db43a-105">Donnez-nous votre avis.</span><span class="sxs-lookup"><span data-stu-id="db43a-105">Give us your feedback</span></span>

<span data-ttu-id="db43a-106">Votre avis compte beaucoup pour nous.</span><span class="sxs-lookup"><span data-stu-id="db43a-106">Your feedback is important to us.</span></span>

* <span data-ttu-id="db43a-107">Si vous avez des questions concernant les documents, [soumettez-les](https://github.com/OfficeDev/office-js-docs-reference/issues) dans ce référentiel.</span><span class="sxs-lookup"><span data-stu-id="db43a-107">To let us know about any questions or issues you find in the docs, [submit an issue](https://github.com/OfficeDev/office-js-docs-reference/issues) in this repository.</span></span> <span data-ttu-id="db43a-108">Assurez-vous de faire état de la version + numéro de build du client que vous utilisez, et fournissez les étapes de repro, la sortie de la console et les messages d’erreur, le cas échéant.</span><span class="sxs-lookup"><span data-stu-id="db43a-108">Make sure you state the version + build number of the client you are using, and provide repro steps, console output, and error messages, as appropriate.</span></span>

* <span data-ttu-id="db43a-109">Vos contributions à cette documentation sont également les bienvenues.</span><span class="sxs-lookup"><span data-stu-id="db43a-109">We also welcome your contributions to this documentation.</span></span> <span data-ttu-id="db43a-110">Pour contribuer, bifurcation de ce référentiel, mettez à jour les fichiers selon vos besoins et envoyez une requête de pull avec vos modifications proposées.</span><span class="sxs-lookup"><span data-stu-id="db43a-110">To contribute, fork this repository, update the files as you deem necessary, and submit a pull request with your proposed changes.</span></span> <span data-ttu-id="db43a-111">Pour plus de détails, consultez la rubrique [Contribuer à cette documentation](Contributing.md).</span><span class="sxs-lookup"><span data-stu-id="db43a-111">For details, see [Contribute to this documentation](Contributing.md).</span></span>

    > <span data-ttu-id="db43a-112">**IMPORTANT**: ne modifiez pas les fichiers dans le [dossier /docs/docs-ref-autogen](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen) de ce référentiel.</span><span class="sxs-lookup"><span data-stu-id="db43a-112">**IMPORTANT**: Do not modify files within the [/docs/docs-ref-autogen](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen) folder of this repository.</span></span> <span data-ttu-id="db43a-113">Tous les fichiers de ce dossier sont créés automatiquement, il n’est donc pas possible de les mettre à jour via une demande de pull.</span><span class="sxs-lookup"><span data-stu-id="db43a-113">All of the files in that folder are autogenerated, so it is not possible to update them via pull request.</span></span> <span data-ttu-id="db43a-114">Pour demander une modification de l’un des fichiers dans le [dossier /docs/docs-ref-autogen,](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen) envoyez un [problème](https://github.com/OfficeDev/office-js-docs-reference/issues) dans ce référentiel.</span><span class="sxs-lookup"><span data-stu-id="db43a-114">To request a change to any of the files in the [/docs/docs-ref-autogen](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen) folder, please [submit an issue](https://github.com/OfficeDev/office-js-docs-reference/issues) in this repository.</span></span> <span data-ttu-id="db43a-115">Vous pouvez en savoir plus sur la façon dont les outils dans ce référentiel [sont ici.](https://github.com/OfficeDev/office-js-docs-reference/blob/master/DocumentationToolingNotes.md)</span><span class="sxs-lookup"><span data-stu-id="db43a-115">You can read more about how the tooling in this repository [here](https://github.com/OfficeDev/office-js-docs-reference/blob/master/DocumentationToolingNotes.md).</span></span>

* <span data-ttu-id="db43a-116">Pour nous faire connaître votre expérience de programmation, ce que vous souhaitez voir dans les versions futures, des exemples de code, etc., entrez vos suggestions et idées sur [Microsoft 365](https://docs.microsoft.com/answers/products/m365)sur Q&R.</span><span class="sxs-lookup"><span data-stu-id="db43a-116">To let us know about your programming experience, what you would like to see in future versions, code samples, and so on, enter your suggestions and ideas at [Microsoft 365 on Q&A](https://docs.microsoft.com/answers/products/m365).</span></span>

## <a name="join-the-microsoft-365-developer-program"></a><span data-ttu-id="db43a-117">Rejoignez le programme Développeur de Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="db43a-117">Join the Microsoft 365 Developer Program</span></span>

<span data-ttu-id="db43a-118">Obtenez un bac à sable, des outils et d’autres ressources gratuits dont vous avez besoin pour créer des solutions pour la plateforme Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="db43a-118">Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.</span></span>

* <span data-ttu-id="db43a-119">[Bac à sable développeur gratuit](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Obtenez un abonnement microsoft 365 E5 développeur gratuit et renouvelable de 90 jours.</span><span class="sxs-lookup"><span data-stu-id="db43a-119">[Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.</span></span>
* <span data-ttu-id="db43a-120">[Packs d’exemples de données](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Configurez automatiquement votre bac à sable en installant les données utilisateur et le contenu pour vous aider à créer vos solutions.</span><span class="sxs-lookup"><span data-stu-id="db43a-120">[Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.</span></span>
* <span data-ttu-id="db43a-121">[Accès aux experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Accédez aux événements de la communauté pour en savoir plus sur les experts Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="db43a-121">[Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.</span></span>
* <span data-ttu-id="db43a-122">[Recommandations personnalisées](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Recherchez rapidement des ressources pour les développeurs à partir de votre tableau de bord personnalisé.</span><span class="sxs-lookup"><span data-stu-id="db43a-122">[Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.</span></span>


## <a name="microsoft-open-source-code-of-conduct"></a><span data-ttu-id="db43a-123">Code de conduite Open Source de Microsoft</span><span class="sxs-lookup"><span data-stu-id="db43a-123">Microsoft Open Source Code of Conduct</span></span>

<span data-ttu-id="db43a-124">Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/).</span><span class="sxs-lookup"><span data-stu-id="db43a-124">This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).</span></span>
<span data-ttu-id="db43a-125">Pour plus d’informations, consultez le FORUM aux questions sur le [code](https://opensource.microsoft.com/codeofconduct/faq/)de conduite ou [contactez opencode@microsoft.com](mailto:opencode@microsoft.com) avec d’autres questions ou commentaires.</span><span class="sxs-lookup"><span data-stu-id="db43a-125">For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/), or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.</span></span>
