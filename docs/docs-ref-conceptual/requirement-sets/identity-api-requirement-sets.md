# <a name="identity-api-requirement-sets"></a>Ensembles de conditions requises de l’API d’identité

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Compléments Office utilisent ensembles spécifiés dans le manifeste ou une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API nécessitant un complément. Pour plus d’informations, voir [définit les versions d’Office et de la spécification](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Compléments Office s’exécuter sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles API d’identité, les applications hôtes Office qui prennent en charge que les numéros exigence set et la version ou la version de l’application Office.

|  Ensemble de conditions requises  | Office 2013 pour Windows | Office 365 pour Windows   |  Office 365 pour iPad  |  Office 365 pour Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com et Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | S/O | Aperçu **& #42 ;** | Bientôt disponible | Aperçu **& #42 ;**| Available | Available| Bientôt disponible | Bientôt disponible |

> **& #42 ;** Pendant la phase d’aperçu, l’API de l’identité est pris en charge sur 2016 Windows et Mac uniquement pour les utilisateurs dans le programme initiés à l’aide de l’option Fast. Pour participer au programme d’initiés, visualiser [Un initié Office](https://products.office.com/office-insider?tab=tab-1). Pour passer à la piste Fast, voir [Initiés Fast](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961).

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

- [Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver le numéro de version et de build pour une application cliente Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- 
  [Présentation d’Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office

Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="identityapi-11"></a>IdentityAPI 1.1 

IdentityAPI 1.1 à connexion unique est la première version de l’API. Pour plus d’informations sur l’API, voir la `getAccessTokenAsync` méthode dans la rubrique de référence [Office.Auth](/javascript/api/office/office.auth) .

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des exigences en matière d’hôtes Office et d’API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
