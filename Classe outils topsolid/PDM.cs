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
    /// Classe utilitaire pour la gestion des opérations PDM (Product Data Management) dans TopSolid.
    /// </summary>
    /// <remarks>
    /// Cette classe fournit des méthodes pour manipuler les projets, dossiers et documents du PDM TopSolid.
    /// Toutes les méthodes gèrent les erreurs en interne et retournent des valeurs par défaut en cas d'échec.
    /// </remarks>
    public static class PDM
    {
        #region Récupération Projet

        /// <summary>
        /// Récupère le PdmObjectId du projet courant.
        /// Si un document est ouvert, utilise son projet parent.
        /// Sinon, utilise le projet actif dans TopSolid.
        /// </summary>
        /// <returns>Le PdmObjectId du projet courant, ou un PdmObjectId vide si aucun projet n'est trouvé.</returns>
        /// <example>
        /// <code>
        /// // Récupérer le projet courant
        /// PdmObjectId projet = PDM.GetCurrentProjectPdmObject();
        /// if (!projet.IsEmpty)
        /// {
        ///     string nomProjet = TSH.Pdm.GetName(projet);
        ///     Console.WriteLine($"Projet actif : {nomProjet}");
        /// }
        /// </code>
        /// </example>
        public static PdmObjectId GetCurrentProjectPdmObject()
        {
            var doc = TSH.Documents.EditedDocument;

            // Si un document est ouvert, récupérer son projet parent
            if (doc != null && doc != DocumentId.Empty)
            {
                try
                {
                    var docPdmObjId = TSH.Documents.GetPdmObject(doc);
                    return TSH.Pdm.GetProject(docPdmObjId);
                }
                catch (Exception)
                {
                    // Si échec, tenter de récupérer le projet courant directement
                }
            }

            // Sinon, récupérer le projet actif dans TopSolid
            try
            {
                return TSH.Pdm.GetCurrentProject();
            }
            catch (Exception)
            {
                return new PdmObjectId(); // Retourne un PdmObjectId vide en cas d'échec
            }
        }

        /// <summary>
        /// Récupère le nom du projet courant.
        /// </summary>
        /// <returns>Le nom du projet, ou une chaîne vide si non trouvé.</returns>
        /// <example>
        /// <code>
        /// // Afficher le nom du projet actif
        /// string nomProjet = PDM.GetCurrentProjectName();
        /// if (!string.IsNullOrEmpty(nomProjet))
        /// {
        ///     MessageBox.Show($"Vous travaillez sur : {nomProjet}");
        /// }
        /// </code>
        /// </example>
        public static string GetCurrentProjectName()
        {
            try
            {
                var projectPdm = GetCurrentProjectPdmObject();
                if (!projectPdm.IsEmpty)
                {
                    return TSH.Pdm.GetName(projectPdm);
                }
            }
            catch (Exception)
            {
                // Silencieux en cas d'erreur
            }

            return string.Empty;
        }

        #endregion

        #region Gestion Documents

        /// <summary>
        /// Récupère le PdmObjectId d'un document donné.
        /// </summary>
        /// <param name="documentId">L'identifiant du document.</param>
        /// <returns>Le PdmObjectId du document, ou un PdmObjectId vide si échec.</returns>
        /// <example>
        /// <code>
        /// // Récupérer le PdmObjectId du document actif
        /// DocumentId docId = TSH.Documents.EditedDocument;
        /// PdmObjectId pdmObj = PDM.GetDocumentPdmObject(docId);
        /// if (!pdmObj.IsEmpty)
        /// {
        ///     Console.WriteLine("Document géré par PDM");
        /// }
        /// </code>
        /// </example>
        public static PdmObjectId GetDocumentPdmObject(DocumentId documentId)
        {
            if (documentId == null || documentId == DocumentId.Empty)
                return new PdmObjectId();

            try
            {
                return TSH.Documents.GetPdmObject(documentId);
            }
            catch (Exception)
            {
                return new PdmObjectId();
            }
        }

        /// <summary>
        /// Vérifie si un document est géré par le PDM.
        /// </summary>
        /// <param name="documentId">L'identifiant du document.</param>
        /// <returns>True si le document est géré par PDM, sinon False.</returns>
        /// <example>
        /// <code>
        /// // Vérifier si le document actif est dans le PDM
        /// DocumentId docId = TSH.Documents.EditedDocument;
        /// if (PDM.IsDocumentInPdm(docId))
        /// {
        ///     MessageBox.Show("Ce document est sous contrôle PDM");
        /// }
        /// else
        /// {
        ///     MessageBox.Show("Document hors PDM");
        /// }
        /// </code>
        /// </example>
        public static bool IsDocumentInPdm(DocumentId documentId)
        {
            var pdmObject = GetDocumentPdmObject(documentId);
            return !pdmObject.IsEmpty;
        }

        /// <summary>
        /// Récupère tous les documents du projet courant de manière récursive.
        /// </summary>
        /// <returns>Liste des PdmObjectId de tous les documents.</returns>
        /// <example>
        /// <code>
        /// // Lister tous les documents du projet
        /// List&lt;PdmObjectId&gt; documents = PDM.GetAllProjectDocuments();
        /// MessageBox.Show($"Nombre total de documents : {documents.Count}");
        /// 
        /// foreach (var doc in documents)
        /// {
        ///     string nomDoc = TSH.Pdm.GetName(doc);
        ///     Console.WriteLine($"- {nomDoc}");
        /// }
        /// </code>
        /// </example>
        public static List<PdmObjectId> GetAllProjectDocuments()
        {
            var allDocuments = new List<PdmObjectId>();

            try
            {
                var projectPdm = GetCurrentProjectPdmObject();
                if (!projectPdm.IsEmpty)
                {
                    // Documents à la racine du projet
                    TSH.Pdm.GetConstituents(projectPdm, out var dossiers, out var documents);
                    
                    if (documents != null && documents.Count > 0)
                    {
                        allDocuments.AddRange(documents);
                    }
                    
                    // Documents dans tous les dossiers
                    if (dossiers != null)
                    {
                        foreach (var dossier in dossiers)
                        {
                            var docsInFolder = GetDocumentsRecursive(dossier);
                            allDocuments.AddRange(docsInFolder);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERREUR : {ex.Message}\n\nStack:\n{ex.StackTrace}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return allDocuments;
        }

        /// <summary>
        /// Récupère récursivement tous les documents d'un dossier et ses sous-dossiers.
        /// </summary>
        /// <param name="folderPdm">Le dossier parent.</param>
        /// <returns>Liste de tous les documents.</returns>
        private static List<PdmObjectId> GetDocumentsRecursive(PdmObjectId folderPdm)
        {
            var allDocuments = new List<PdmObjectId>();

            try
            {
                TSH.Pdm.GetConstituents(folderPdm, out var subFolders, out var documents);
                
                // Ajouter les documents du dossier courant
                if (documents != null && documents.Count > 0)
                {
                    allDocuments.AddRange(documents);
                }
                
                // Récursion dans les sous-dossiers
                if (subFolders != null && subFolders.Count > 0)
                {
                    foreach (var subFolder in subFolders)
                    {
                        var deeperDocuments = GetDocumentsRecursive(subFolder);
                        allDocuments.AddRange(deeperDocuments);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERREUR dans GetDocumentsRecursive : {ex.Message}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return allDocuments;
        }

        #endregion

        #region Gestion Dossiers (à implémenter)

        /// <summary>
        /// Récupère la liste des dossiers du projet courant de manière récursive.
        /// </summary>
        /// <returns>Liste des PdmObjectId des dossiers.</returns>
        /// <example>
        /// <code>
        /// // Lister tous les dossiers du projet
        /// List&lt;PdmObjectId&gt; dossiers = PDM.GetProjectFolders();
        /// foreach (var dossier in dossiers)
        /// {
        ///     string nomDossier = TSH.Pdm.GetName(dossier);
        ///     Console.WriteLine($"Dossier : {nomDossier}");
        /// }
        /// </code>
        /// </example>
        public static List<PdmObjectId> GetProjectFolders()
        {
            var folders = new List<PdmObjectId>();

            try
            {
                var projectPdm = GetCurrentProjectPdmObject();
                if (!projectPdm.IsEmpty)
                {
                    TSH.Pdm.GetConstituents(projectPdm, out var dossiers, out var documents);
                    
                    if (dossiers != null)
                    {
                        folders.AddRange(dossiers);
                        
                        // Récupération récursive des sous-dossiers
                        foreach (var dossier in dossiers)
                        {
                            var subFolders = GetSubFoldersRecursive(dossier);
                            folders.AddRange(subFolders);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERREUR : {ex.Message}\n\nStack:\n{ex.StackTrace}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return folders;
        }

        /// <summary>
        /// Récupère récursivement tous les sous-dossiers d'un dossier donné.
        /// </summary>
        /// <param name="folderPdm">Le dossier parent.</param>
        /// <returns>Liste des sous-dossiers à tous les niveaux.</returns>
        private static List<PdmObjectId> GetSubFoldersRecursive(PdmObjectId folderPdm)
        {
            var allSubFolders = new List<PdmObjectId>();

            try
            {
                TSH.Pdm.GetConstituents(folderPdm, out var subFolders, out var files);
                
                if (subFolders != null && subFolders.Count > 0)
                {
                    allSubFolders.AddRange(subFolders);
                    
                    // Appel récursif pour chaque sous-dossier
                    foreach (var subFolder in subFolders)
                    {
                        var deeperFolders = GetSubFoldersRecursive(subFolder);
                        allSubFolders.AddRange(deeperFolders);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERREUR dans GetSubFoldersRecursive : {ex.Message}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return allSubFolders;
        }

        /// <summary>
        /// Récupère la liste des fichiers dans un dossier PDM donné (non récursif).
        /// </summary>
        /// <param name="folderPdmObject">Le PdmObjectId du dossier.</param>
        /// <returns>Liste des PdmObjectId des fichiers du dossier uniquement.</returns>
        /// <example>
        /// <code>
        /// // Récupérer les fichiers d'un dossier spécifique
        /// List&lt;PdmObjectId&gt; dossiers = PDM.GetProjectFolders();
        /// if (dossiers.Count > 0)
        /// {
        ///     PdmObjectId premierDossier = dossiers[0];
        ///     List&lt;PdmObjectId&gt; fichiers = PDM.GetFilesInFolder(premierDossier);
        ///     MessageBox.Show($"Nombre de fichiers : {fichiers.Count}");
        /// }
        /// </code>
        /// </example>
        public static List<PdmObjectId> GetFilesInFolder(PdmObjectId folderPdmObject)
        {
            var files = new List<PdmObjectId>();

            try
            {
                if (!folderPdmObject.IsEmpty)
                {
                    TSH.Pdm.GetConstituents(folderPdmObject, out var subFolders, out var documents);
                    
                    if (documents != null && documents.Count > 0)
                    {
                        files.AddRange(documents);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERREUR : {ex.Message}\n\nStack:\n{ex.StackTrace}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return files;
        }

        #endregion

        #region Utilitaires

        /// <summary>
        /// Récupère le nom d'un objet PDM.
        /// </summary>
        /// <param name="pdmObject">L'objet PDM.</param>
        /// <returns>Le nom de l'objet, ou une chaîne vide si échec.</returns>
        /// <example>
        /// <code>
        /// // Récupérer le nom d'un objet PDM
        /// PdmObjectId projet = PDM.GetCurrentProjectPdmObject();
        /// string nom = PDM.GetName(projet);
        /// Console.WriteLine($"Nom : {nom}");
        /// </code>
        /// </example>
        public static string GetName(PdmObjectId pdmObject)
        {
            try
            {
                if (!pdmObject.IsEmpty)
                {
                    return TSH.Pdm.GetName(pdmObject);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERREUR : {ex.Message}\n\nStack:\n{ex.StackTrace}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return string.Empty;
        }

        #endregion
    }

    
}
