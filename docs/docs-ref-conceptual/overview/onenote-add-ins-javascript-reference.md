# <a name="onenote-javascript-api-overview"></a>Vue d’ensemble de l’API JavaScript de OneNote

S’applique à : OneNote Online

Les liens suivants affichent les objets OneNote de niveau supérieur disponibles dans l’API. Liaison d’objet de chaque page contient une description des événements, propriétés et méthodes disponibles sur l’objet. Cliquez sur ces liens pour en savoir plus. 
    
- [Application](/javascript/api/onenote/onenote.application) : Objet de niveau supérieur utilisé pour accéder à tous les objets OneNote globalement adressables, tels que le bloc-notes actif et la section active.

- [Bloc-notes](/javascript/api/onenote/onenote.notebook) : Bloc-notes. Les blocs-notes contiennent des groupes de sections et des sections.
    - [NotebookCollection](/javascript/api/onenote/onenote.notebookcollection) : Collection de blocs-notes.

- [SectionGroup](/javascript/api/onenote/onenote.sectiongroup) : Groupe de sections. Les groupes de sections contiennent des sections et des groupes de sections.
    - [SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection) : Collection de groupes de sections.

- [Section](/javascript/api/onenote/onenote.section) : Section. Les sections contiennent des pages.
    - [SectionCollection](/javascript/api/onenote/onenote.sectioncollection) : Collection de sections.

- [Page](/javascript/api/onenote/onenote.page) : Page. Les pages contiennent des objets PageContent.
    - [PageCollection](/javascript/api/onenote/onenote.pagecollection) : Collection de pages.

- [PageContent](/javascript/api/onenote/onenote.pagecontent) : Zone de niveau supérieur sur une page qui contient des types de contenu tels que des plans ou des images. Un objet PageContent peut être affecté à une position sur la page.
    - [PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection) : Collection d’objets PageContent qui représente le contenu d’une page.

- [Outline](/javascript/api/onenote/onenote.outline) : Conteneur pour les objets Paragraph. Un plan est un enfant direct d’un objet PageContent.

- [Image](/javascript/api/onenote/onenote.image) : Objet Image. Une image peut être un enfant direct d’un objet Paragraph ou PageContent.

- [Paragraph](/javascript/api/onenote/onenote.paragraph) : Conteneur pour le contenu visible d’une page. Un paragraphe est un enfant direct d’un plan.
    - [ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection) : Collection d’objets Paragraph dans un plan.

- [Richtext](/javascript/api/onenote/onenote.richtext) : Objet RichText.

- [Table](/javascript/api/onenote/onenote.table) : Conteneur pour les objets TableRow.

- [TableRow](/javascript/api/onenote/onenote.tablerow) : Conteneur pour les objets TableCell.
    - [TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection) : Collection d’objets TableRow dans un tableau.
 
- [TableCell](/javascript/api/onenote/onenote.tablecell) : Conteneur pour les objets Paragraph.
    - [TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection) : Collection d’objets TableCell dans un élément TableRow.

## <a name="onenote-javascript-api-reference"></a>Référence de l’API JavaScript de OneNote

Pour plus d’informations sur OneNote JavaScript API, consultez la [documentation de référence API JavaScript de OneNote](/javascript/api/onenote).

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la programmation de l’API JavaScript de OneNote](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [Créer votre premier complément OneNote](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-getting-started)
- [Exemple de grille d’évaluation](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Vue d’ensemble de la plateforme des compléments Office](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
