# <a name="outlook-add-in-api-requirement-set-16"></a>Outlook complément nécessite API définie 1.6

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

## <a name="whats-new-in-16"></a>Quelles sont les nouveautés dans 1.6 ?

Ensemble de conditions requises 1.6 inclut toutes les fonctionnalités de [spécification ensemble de 1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md). Il ajouté les fonctionnalités suivantes.

- Ajout nouvelles API pour les compléments contextuelles pour obtenir de l’entité ou RegEx correspond à que l’utilisateur sélectionné pour activer le complément.
- Ajout d’une nouvelle API pour ouvrir un nouveau formulaire de message.
- Ajout de la possibilité pour le complément déterminer le type de compte de boîte aux lettres de l’utilisateur.

### <a name="change-log"></a>Journal des modifications

- Ajouté [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities): ajoute une nouvelle fonction qui obtient les entités trouvées dans une correspondance en surbrillance un utilisateur a sélectionné. Les correspondances en surbrillance s’appliquent aux compléments contextuels.
- Ajouté [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object): ajoute une nouvelle fonction qui renvoie les valeurs de chaîne dans une correspondance en surbrillance qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux compléments contextuels.
- Ajouté [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters): ajoute une nouvelle fonction qui ouvre un nouveau formulaire de message.
- Ajouté [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): ajoute un nouveau membre au profil d’utilisateur qui indique le type de compte de l’utilisateur.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](https://docs.microsoft.com/outlook/add-ins/quick-start)