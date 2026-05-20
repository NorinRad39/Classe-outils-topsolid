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
        /// Constructeur privé pour créer une instance de <see cref="ProjetPDM"/>.
        /// </summary>
        /// <param name="pdmObjectId">L'identifiant PDM du projet. Doit être un <see cref="PdmObjectId"/> valide (non vide).</param>
        /// <remarks>
        /// Cette instance initialise les champs internes nécessaires :
        /// - stocke l'identifiant PDM dans <see cref="_pdmObjectId"/>.
        /// - récupère et affecte le nom du projet via <c>TSH.Pdm.GetName(pdmObjectId)</c>.
        ///
        /// Le constructeur est privé afin de forcer la création d'instances via les
        /// méthodes/factories exposées par la classe (par exemple <see cref="ProjetCourant"/> et <see cref="GetLibraries"/>).
        /// </remarks>
        /// <exception cref="ArgumentException">Si <paramref name="pdmObjectId"/> est vide (vérifier avec <c>pdmObjectId.IsEmpty</c> avant l'appel).</exception>
        /// <exception cref="Exception">
        /// Peut lever une exception si l'appel à l'API TopSolid (<c>TSH.Pdm.GetName</c>) échoue.
        /// Gérez les exceptions appelantes si nécessaire.
        /// </exception>
        /// <example>
        /// <code>
        /// // Exemple d'usage indirect : obtenir le projet courant
        /// var projet = ProjetPDM.ProjetCourant;
        /// // -> n'appelez pas directement le constructeur depuis l'extérieur (il est privé)
        /// </code>
        /// </example>
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

                /// <summary>
        /// Récupère la liste des projets PDM considérés comme des bibliothèques.
        /// </summary>
        /// <returns>
        /// Une <see cref="List{ProjetPDM}"/> contenant une instance <see cref="ProjetPDM"/> pour chaque projet bibliothèque trouvé.
        /// Si aucun projet n'est trouvé ou si l'appel à l'API PDM retourne <c>null</c>, une liste vide est retournée.
        /// </returns>
        /// <remarks>
        /// Cette méthode appelle <c>TSH.Pdm.GetProjects(false, true)</c> pour obtenir les projets de type bibliothèque.
        /// Les entrées renvoyées par l'API PDM sont filtrées : seuls les <see cref="PdmObjectId"/> non vides sont transformés
        /// en instances <see cref="ProjetPDM"/>.
        /// La méthode ne lève pas d'exception si l'API retourne <c>null</c> ; elle renvoie simplement une liste vide.
        /// </remarks>
        /// <example>
        /// <code>
        /// // Obtenir toutes les bibliothèques et afficher leur nom
        /// var biblios = ProjetPDM.GetLibraries();
        /// foreach (var biblio in biblios)
        /// {
        ///     Console.WriteLine(biblio.Nom);
        /// }
        /// </code>
        /// </example>
        public static List<ProjetPDM> GetLibraries()
        {
            var result = new List<ProjetPDM>();
            var allProject = TSH.Pdm.GetProjects(false, true);
            if (allProject == null)
                return result;

            foreach (var project in allProject)
            {
                if (!project.IsEmpty)
                    result.Add(new ProjetPDM(project));
            }

            return result;
        }

       

        /// <summary>
        /// Récupère les noms des bibliothèques PDM.
        /// </summary>
        /// <returns>
        /// Une <see cref="List{String}"/> contenant les noms des projets bibliothèques.
        /// Si aucun projet n'est trouvé ou si l'appel à l'API PDM retourne <c>null</c>, une liste vide est retournée.
        /// </returns>
        /// <remarks>
        /// Cette méthode demande la liste des projets bibliothèques via <c>TSH.Pdm.GetProjects(false, true)</c>,
        /// puis récupère le nom de chaque projet avec <c>TSH.Pdm.GetName</c>.
        /// Les erreurs lors de la lecture du nom d'un projet sont ignorées (bloc <c>try/catch</c> vide),
        /// ce qui permet de continuer la collecte des noms même si certains projets sont invalides ou inaccessibles.
        /// Notez que des doublons sont possibles si plusieurs projets partagent le même nom.
        /// </remarks>
        /// <example>
        /// <code>
        /// // Obtenir et afficher les noms des bibliothèques
        /// var noms = ProjetPDM.GetLibraryNames();
        /// Console.WriteLine($"Bibliothèques trouvées : {noms.Count}");
        /// foreach (var nom in noms)
        /// {
        ///     Console.WriteLine($"- {nom}");
        /// }
        /// </code>
        /// </example>
        public static List<string> GetLibraryNames()
        {
            var result = new List<string>();
            var allProject = TSH.Pdm.GetProjects(false, true);
            if (allProject == null)
                return result;

            foreach (var project in allProject)
            {
                try
                {
                    result.Add(TSH.Pdm.GetName(project));
                }
                catch
                {
                    // Ignore les projets invalides
                }
            }

            return result;
        }

        /// <summary>
        /// Accès en tant qu'objet à la collection des bibliothèques PDM.
        /// </summary>
        /// <remarks>
        /// Equivalent à appeler <see cref="GetLibraries"/> ; retourne une collection en lecture seule.
        /// Utile pour écrire du code plus expressif : <c>var libs = ProjetPDM.Libraries;</c>
        /// </remarks>
        /// <example>
        /// <code>
        /// // Récupérer toutes les bibliothèques
        /// var biblios = ProjetPDM.Libraries;
        ///
        /// // Récupérer la première bibliothèque
        /// var premiere = biblios.FirstOrDefault();
        ///
        /// // Trouver une bibliothèque par nom
        /// var myBiblio = ProjetPDM.GetLibraryByName("MaBibliotheque");
        /// if (myBiblio != null)
        /// {
        ///     Console.WriteLine(myBiblio.Nom);
        /// }
        /// </code>
        /// </example>
        public static IReadOnlyList<ProjetPDM> Libraries => GetLibraries();

        /// <summary>
        /// Recherche une bibliothèque PDM par son nom.
        /// </summary>
        /// <param name="name">Nom exact de la bibliothèque recherchée.</param>
        /// <param name="ignoreCase">Si <c>true</c>, la recherche ignore la casse (par défaut <c>true</c>).</param>
        /// <returns>L'instance <see cref="ProjetPDM"/> correspondante si trouvée, sinon <c>null</c>.</returns>
        /// <remarks>
        /// Effectue une recherche en mémoire sur la liste retournée par <see cref="GetLibraries"/>.
        /// </remarks>
               /// <summary>
        /// Recherche une bibliothèque PDM par son nom.
        /// </summary>
        /// <param name="name">Nom exact de la bibliothèque recherchée. La valeur <c>null</c> ou chaîne vide retourne <c>null</c>.</param>
        /// <param name="ignoreCase">Si <c>true</c>, la comparaison ignore la casse (par défaut : <c>true</c>).</param>
        /// <returns>
        /// L'instance <see cref="ProjetPDM"/> correspondant au nom si trouvée ; sinon <c>null</c>.
        /// </returns>
        /// <remarks>
        /// - La recherche est effectuée en mémoire sur la collection renvoyée par <see cref="GetLibraries"/>. 
        /// - Cette méthode effectue un appel vers l'API PDM (via <see cref="GetLibraries"/>) à chaque invocation : pour des recherches répétées,
        ///   préférez récupérer la collection via la propriété <see cref="Libraries"/> puis interroger cette collection localement afin d'éviter
        ///   des appels réseau/IO supplémentaires.
        /// - Si <paramref name="name"/> est <c>null</c> ou vide, la méthode retourne immédiatement <c>null</c>.
        /// </remarks>
        /// <example>
        /// <code>
        /// // Recherche insensible à la casse (par défaut)
        /// var biblio = ProjetPDM.GetLibraryByName("MaBibliotheque");
        /// if (biblio != null)
        /// {
        ///     Console.WriteLine($"Bibliothèque trouvée : {biblio.Nom}");
        /// }
        /// 
        /// // Recherche sensible à la casse
        /// var biblioExacte = ProjetPDM.GetLibraryByName("MaBibliotheque", ignoreCase: false);
        /// </code>
        /// </example>
        public static ProjetPDM GetLibraryByName(string name, bool ignoreCase = true)
        {
            if (string.IsNullOrEmpty(name))
                return null;

            var comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            return GetLibraries().FirstOrDefault(p => string.Equals(p.Nom, name, comparison));
        }

        /// <summary>
        /// Tente de récupérer une bibliothèque PDM par son nom.
        /// </summary>
        /// <param name="name">Nom recherché de la bibliothèque. Doit être non null et non vide.</param>
        /// <param name="projet">Sortie : l'instance <see cref="ProjetPDM"/> correspondant au nom, ou <c>null</c> si non trouvée.</param>
        /// <param name="ignoreCase">Si <c>true</c>, la comparaison ignore la casse (par défaut : <c>true</c>).</param>
        /// <returns><c>true</c> si une bibliothèque correspondant au nom a été trouvée ; sinon <c>false</c>.</returns>
        /// <remarks>
        /// Cette méthode délègue la recherche à <see cref="GetLibraryByName(string,bool)"/> qui effectue la recherche
        /// en mémoire sur la collection retournée par <see cref="GetLibraries"/>. La méthode est sûre vis-à-vis des
        /// entrées nulles ou vides : elle renverra <c>false</c> et affectera <paramref name="projet"/> à <c>null</c>.
        /// </remarks>
        /// <example>
        /// <code>
        /// if (ProjetPDM.TryGetLibrary("MaBibliotheque", out var biblio))
        /// {
        ///     Console.WriteLine($"Bibliothèque trouvée : {biblio.Nom}");
        /// }
        /// else
        /// {
        ///     Console.WriteLine("Bibliothèque introuvable.");
        /// }
        /// </code>
        /// </example>
        public static bool TryGetLibrary(string name, out ProjetPDM projet, bool ignoreCase = true)
        {
            if (string.IsNullOrEmpty(name))
            {
                projet = null;
                return false;
            }

            projet = GetLibraryByName(name, ignoreCase);
           
            return projet != null;
        }
    }
}
