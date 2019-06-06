Voici un script Powershell permettant de: 
1.	La création d’un groupe Teams
2.	L’activation du Plan par défaut du groupe Office365

Pour ce faire, il vous faudra (avec des droits d'administration):
*	Installer le module Powershell Microsoft Teams depuis la Powershell Gallery (v1.0.0 au moment de l’édition du script)
```Powershell
Install-Module -Name MicrosoftTeams
```
*	Installer le module Powershell AzureAD
```Powershell
Install-Module -Name AzureAD
```
*	Exécuter le script avec un profil utilisateur ayant les droits d’Administration sur le Tenant AAD cible

Voici des liens qui vous détaillerons les prérequis pour la bonne exécution du POC : 
1.	[Blog Technet](https://blogs.technet.microsoft.com/skypehybridguy/2017/11/07/microsoft-teams-powershell-support/) Installation et usage des Modules Microsoft Teams pour Powershell 
2.	[Blog Technet](https://blogs.technet.microsoft.com/cloudlojik/2018/06/29/connecting-to-microsoft-graph-with-a-native-app-using-powershell/) Se connecter à l’API Graph au travers d’une application native :
a.	Il est important de penser à sélectionner le rôle Group.ReadWriteAll pour Microsoft Graph lors de la création de l’application sur AAD

Enfin, voici des ressources supplémentaires afin d’approfondir la question : 
*	[Github](https://github.com/MicrosoftDocs/office-docs-powershell/blob/master/teams/teams-ps/teams/New-Team.md) Documentation pour l’usage des différentes CmdLet PS Microsoft Teams
*	[Developers](https://developer.microsoft.com/en-us/graph/graph-explorer) Graph Explorer vous permettra de tester des requêtes REST sur l’API MSGraph
*	[MS Docs](https://docs.microsoft.com/en-us/graph/api/planner-post-plans?view=graph-rest-1.0&tabs=cs#permissions) Documentation de l’API Graph (Article : CreatePlannerPlan)

