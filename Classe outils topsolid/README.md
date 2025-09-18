# OutilsTs

Bibliothèque .NET pour gérer et automatiser les données des documents dans TopSolid.

## Présentation

OutilsTs facilite l’accès aux propriétés, paramètres, opérations et métadonnées des documents TopSolid.  
Compatible avec .NET Framework 4.8 et utilisable dans des applications WPF ou WinForms (x64).

## Prérequis et compatibilité

- .NET Framework 4.8  
- Windows x64 (le projet cible explicitement x64)  
- TopSolid 7.19 installé (API disponibles et licenciées)  
- Exécution dans un contexte TopSolid (hôte/add-in) recommandée  

## Installation

Depuis la Console du Gestionnaire de packages (Visual Studio) :

```powershell
Install-Package OutilsTs
```

Via la CLI .NET :

```powershell
dotnet add package OutilsTs
```

## Utilisation

### Récupérer des informations sur le document courant

```csharp
using OutilsTs;

var doc = new Document();

// Nom du document
string nom = doc.DocNomTxt;

// Type/extension PDM
string ext = doc.DocExtention;

// Statuts
bool estDerive = doc.DocDerived;
bool estElectrode = doc.DocIsElectrode;

// Numéro d’OP (si présent)
string op = doc.OP;
```

### Récupérer les opérations CAM (si document CAM)

```csharp
using OutilsTs;

var doc = new Document();

bool estCam = doc.DocuCam;
var operationsCam = doc.CamOperations; // Liste d'ElementId (TopSolid)
```

### Afficher un message (WPF/WinForms)

```csharp
using OutilsTs;

UiHelper.ShowMessage("Opération terminée", "Information", MessageType.Info);
// MessageType.Warning et MessageType.Error sont également disponibles
```

## Notes d’exécution

- Les API TopSolid (TSH/TSCH/TSHD) nécessitent un environnement TopSolid actif.  
- En cas d’exception, un message utilisateur peut s’afficher et une trace est écrite dans `errors.log` (répertoire courant du processus).  

## Journalisation

Les erreurs capturées sont ajoutées à `errors.log` avec un horodatage basique.

## Licence

MIT

## Liens

- NuGet : [https://www.nuget.org/packages/OutilsTs](https://www.nuget.org/packages/OutilsTs)  
- Repository : [https://github.com/NorinRad39/Classe-outils-topsolid](https://github.com/NorinRad39/Classe-outils-topsolid)  

## Historique

- 1.0.0 — Première version publique
