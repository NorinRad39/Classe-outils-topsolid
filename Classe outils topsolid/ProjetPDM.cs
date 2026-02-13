using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TopSolid.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;

namespace OutilsTs
{
    /// <summary>
    /// Représente un projet PDM TopSolid avec accès simplifié à ses constituants.
    /// </summary>
    /// <example>
    /// <code>
    /// var projet = ProjetPDM.ProjetCourant;
    /// if (projet != null)
    /// {
    ///     Console.WriteLine($"Projet : {projet.Nom}");
    ///     Console.WriteLine($"Documents : {projet.Documents.Count}");
    /// }
    /// </code>
    /// </example>
    public class ProjetPDM
    {
        private PdmObjectId _pdmObjectId;
        private List<PdmObjectId> _documents;
        private List<PdmObjectId> _dossiers;

        /// <summary>
        /// Obtient le PdmObjectId du projet.
        /// </summary>
        /// <remarks>
        /// Utilisez cette propriété pour accéder directement à l'identifiant PDM si vous avez besoin
        /// d'appeler des méthodes de l'API TopSolid qui nécessitent un PdmObjectId.
        /// </remarks>
        /// <example>
        /// <code>
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet != null)
        /// {
        ///     PdmObjectId pdmId = projet.PdmObjectId;
        ///     string chemin = TSH.Pdm.GetPath(pdmId);
        ///     Console.WriteLine($"Chemin du projet : {chemin}");
        /// }
        /// </code>
        /// </example>
        public PdmObjectId PdmObjectId => _pdmObjectId;

        /// <summary>
        /// Obtient le nom du projet.
        /// </summary>
        /// <remarks>
        /// Le nom est récupéré lors de la création de l'instance et peut être rafraîchi avec la méthode Rafraichir().
        /// </remarks>
        /// <example>
        /// <code>
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet != null)
        /// {
        ///     MessageBox.Show($"Vous travaillez sur le projet : {projet.Nom}", 
        ///                     "Information", 
        ///                     MessageBoxButtons.OK, 
        ///                     MessageBoxIcon.Information);
        /// }
        /// </code>
        /// </example>
        public string Nom { get; private set; }

        /// <summary>
        /// Obtient tous les documents du projet de manière récursive (chargement à la demande).
        /// </summary>
        /// <remarks>
        /// Les documents sont chargés la première fois que la propriété est accédée.
        /// Appeler Rafraichir() pour recharger les données si le contenu du projet a changé.
        /// Inclut tous les documents de tous les dossiers et sous-dossiers.
        /// </remarks>
        /// <example>
        /// <code>
        /// // Exemple 1 : Lister tous les documents
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet != null)
        /// {
        ///     Console.WriteLine($"Total documents : {projet.Documents.Count}");
        ///     foreach (var doc in projet.Documents)
        ///     {
        ///         Console.WriteLine($"- {TSH.Pdm.GetName(doc)}");
        ///     }
        /// }
        /// 
        /// // Exemple 2 : Compter les documents par type
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet != null)
        /// {
        ///     int nbPieces = 0;
        ///     int nbAssemblages = 0;
        ///     
        ///     foreach (var doc in projet.Documents)
        ///     {
        ///         string nom = TSH.Pdm.GetName(doc);
        ///         if (nom.Contains("_Part")) nbPieces++;
        ///         if (nom.Contains("_Assembly")) nbAssemblages++;
        ///     }
        ///     
        ///     Console.WriteLine($"Pièces : {nbPieces}, Assemblages : {nbAssemblages}");
        /// }
        /// 
        /// // Exemple 3 : Ouvrir le premier document trouvé
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet != null &amp;&amp; projet.Documents.Count > 0)
        /// {
        ///     PdmObjectId premierDoc = projet.Documents[0];
        ///     DocumentId docId = TSH.Documents.GetDocument(premierDoc);
        ///     TSH.Documents.Open(docId);
        /// }
        /// </code>
        /// </example>
        public List<PdmObjectId> Documents
        {
            get
            {
                if (_documents == null)
                {
                    _documents = PDM.GetAllProjectDocuments();
                }
                return _documents;
            }
        }

        /// <summary>
        /// Obtient tous les dossiers du projet de manière récursive (chargement à la demande).
        /// </summary>
        /// <remarks>
        /// Les dossiers sont chargés la première fois que la propriété est accédée.
        /// Appeler Rafraichir() pour recharger les données si la structure du projet a changé.
        /// Inclut tous les dossiers et sous-dossiers à tous les niveaux.
        /// </remarks>
        /// <example>
        /// <code>
        /// // Exemple 1 : Lister tous les dossiers
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet != null)
        /// {
        ///     Console.WriteLine($"Total dossiers : {projet.Dossiers.Count}");
        ///     foreach (var dossier in projet.Dossiers)
        ///     {
        ///         Console.WriteLine($"- {TSH.Pdm.GetName(dossier)}");
        ///     }
        /// }
        /// 
        /// // Exemple 2 : Chercher un dossier spécifique
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet != null)
        /// {
        ///     foreach (var dossier in projet.Dossiers)
        ///     {
        ///         string nomDossier = TSH.Pdm.GetName(dossier);
        ///         if (nomDossier.Contains("Electrodes"))
        ///         {
        ///             Console.WriteLine($"Dossier électrodes trouvé : {nomDossier}");
        ///             
        ///             // Récupérer les fichiers de ce dossier
        ///             var fichiers = PDM.GetFilesInFolder(dossier);
        ///             Console.WriteLine($"  Contient {fichiers.Count} fichiers");
        ///         }
        ///     }
        /// }
        /// </code>
        /// </example>
        public List<PdmObjectId> Dossiers
        {
            get
            {
                if (_dossiers == null)
                {
                    _dossiers = PDM.GetProjectFolders();
                }
                return _dossiers;
            }
        }

        /// <summary>
        /// Constructeur privé.
        /// </summary>
        /// <param name="pdmObjectId">L'identifiant PDM du projet.</param>
        private ProjetPDM(PdmObjectId pdmObjectId)
        {
            _pdmObjectId = pdmObjectId;
            Nom = TSH.Pdm.GetName(pdmObjectId);
        }

        /// <summary>
        /// Obtient le projet PDM courant.
        /// </summary>
        /// <returns>Une instance de ProjetPDM représentant le projet actif, ou null si aucun projet n'est ouvert.</returns>
        /// <remarks>
        /// Cette propriété tente de récupérer le projet du document actif, sinon le projet actif dans TopSolid.
        /// Vérifiez toujours que le résultat n'est pas null avant de l'utiliser.
        /// </remarks>
        /// <example>
        /// <code>
        /// // Exemple 1 : Vérification basique
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet != null)
        /// {
        ///     Console.WriteLine($"Projet actif : {projet.Nom}");
        /// }
        /// else
        /// {
        ///     MessageBox.Show("Aucun projet TopSolid ouvert !", "Attention");
        /// }
        /// 
        /// // Exemple 2 : Utilisation avec vérification de validité
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet != null &amp;&amp; projet.EstValide)
        /// {
        ///     // Traiter le projet en toute sécurité
        ///     foreach (var doc in projet.Documents)
        ///     {
        ///         // ...
        ///     }
        /// }
        /// 
        /// // Exemple 3 : Pattern complet avec try-catch
        /// try
        /// {
        ///     var projet = ProjetPDM.ProjetCourant;
        ///     if (projet == null)
        ///     {
        ///         MessageBox.Show("Veuillez ouvrir un projet TopSolid");
        ///         return;
        ///     }
        ///     
        ///     // Votre code ici
        ///     Console.WriteLine($"Traitement du projet {projet.Nom}");
        /// }
        /// catch (Exception ex)
        /// {
        ///     MessageBox.Show($"Erreur : {ex.Message}");
        /// }
        /// </code>
        /// </example>
        public static ProjetPDM ProjetCourant
        {
            get
            {
                var pdmObj = PDM.GetCurrentProjectPdmObject();
                if (pdmObj.IsEmpty)
                    return null;

                return new ProjetPDM(pdmObj);
            }
        }

        /// <summary>
        /// Rafraîchit les données du projet (recharge documents et dossiers).
        /// </summary>
        /// <remarks>
        /// Appelez cette méthode après avoir ajouté/supprimé des documents ou dossiers dans TopSolid
        /// pour obtenir les données à jour. Les listes Documents et Dossiers seront rechargées
        /// lors du prochain accès.
        /// </remarks>
        /// <example>
        /// <code>
        /// // Exemple 1 : Rafraîchissement basique
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet != null)
        /// {
        ///     int nbDocAvant = projet.Documents.Count;
        ///     Console.WriteLine($"Documents avant : {nbDocAvant}");
        ///     
        ///     // ... créer de nouveaux documents dans TopSolid ...
        ///     
        ///     projet.Rafraichir();
        ///     int nbDocApres = projet.Documents.Count;
        ///     Console.WriteLine($"Documents après : {nbDocApres}");
        /// }
        /// 
        /// // Exemple 2 : Surveiller les changements
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet != null)
        /// {
        ///     while (true) // Boucle de traitement
        ///     {
        ///         // Traiter les documents
        ///         foreach (var doc in projet.Documents)
        ///         {
        ///             // ...
        ///         }
        ///         
        ///         // Rafraîchir pour la prochaine itération
        ///         projet.Rafraichir();
        ///         
        ///         System.Threading.Thread.Sleep(5000); // Attendre 5 secondes
        ///     }
        /// }
        /// 
        /// // Exemple 3 : Rafraîchir après modification du nom
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet != null)
        /// {
        ///     Console.WriteLine($"Nom actuel : {projet.Nom}");
        ///     
        ///     // ... renommer le projet dans TopSolid ...
        ///     
        ///     projet.Rafraichir();
        ///     Console.WriteLine($"Nouveau nom : {projet.Nom}");
        /// }
        /// </code>
        /// </example>
        public void Rafraichir()
        {
            _documents = null;
            _dossiers = null;
            Nom = TSH.Pdm.GetName(_pdmObjectId);
        }

        /// <summary>
        /// Vérifie si le projet est valide (son PdmObjectId n'est pas vide).
        /// </summary>
        /// <remarks>
        /// Utilisez cette propriété pour vérifier que le projet est toujours accessible
        /// avant d'effectuer des opérations.
        /// </remarks>
        /// <example>
        /// <code>
        /// // Exemple 1 : Vérification simple
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet != null &amp;&amp; projet.EstValide)
        /// {
        ///     Console.WriteLine("Projet valide et accessible");
        /// }
        /// else
        /// {
        ///     Console.WriteLine("Projet invalide ou inaccessible");
        /// }
        /// 
        /// // Exemple 2 : Pattern de vérification complet
        /// var projet = ProjetPDM.ProjetCourant;
        /// if (projet == null)
        /// {
        ///     MessageBox.Show("Aucun projet ouvert");
        ///     return;
        /// }
        /// 
        /// if (!projet.EstValide)
        /// {
        ///     MessageBox.Show("Le projet n'est pas valide");
        ///     return;
        /// }
        /// 
        /// // Le projet est OK, continuer le traitement
        /// Console.WriteLine($"Traitement de {projet.Nom}");
        /// 
        /// // Exemple 3 : Utilisation dans une condition
        /// var projet = ProjetPDM.ProjetCourant;
        /// string message = (projet != null &amp;&amp; projet.EstValide) 
        ///     ? $"Projet OK : {projet.Nom}" 
        ///     : "Pas de projet valide";
        /// MessageBox.Show(message);
        /// </code>
        /// </example>
        public bool EstValide => !_pdmObjectId.IsEmpty;
    }
}
