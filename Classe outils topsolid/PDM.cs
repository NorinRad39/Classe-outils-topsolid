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
    /// Classe utilitaire statique pour la gestion des opérations PDM (Product Data Management) dans TopSolid.
    /// Couvre les projets, bibliothèques, dossiers et documents.
    /// </summary>
    /// <remarks>
    /// Toutes les méthodes gèrent les erreurs en interne et retournent des valeurs par défaut en cas d'échec.
    /// Les méthodes acceptant un <see cref="PdmObjectId"/> opèrent sur l'objet fourni.
    /// Les surcharges sans paramètre opèrent sur le projet courant.
    /// </remarks>
    public static class PDM
    {
        #region Projet courant

        /// <summary>
        /// Récupère le <see cref="PdmObjectId"/> du projet courant.
        /// Si un document est ouvert, utilise son projet parent. Sinon, utilise le projet actif dans TopSolid.
        /// </summary>
        /// <returns>Le <see cref="PdmObjectId"/> du projet courant, ou <see cref="PdmObjectId.Empty"/> si aucun projet n'est trouvé.</returns>
        /// <example>
        /// <code>
        /// PdmObjectId projet = PDM.GetCurrentProjectPdmObject();
        /// if (!projet.IsEmpty)
        ///     Console.WriteLine(PDM.GetName(projet));
        /// </code>
        /// </example>
        public static PdmObjectId GetCurrentProjectPdmObject()
        {
            var doc = TSH.Documents.EditedDocument;

            if (doc != null && doc != DocumentId.Empty)
            {
                try
                {
                    var docPdmObjId = TSH.Documents.GetPdmObject(doc);
                    return TSH.Pdm.GetProject(docPdmObjId);
                }
                catch (Exception) { }
            }

            try
            {
                return TSH.Pdm.GetCurrentProject();
            }
            catch (Exception)
            {
                return new PdmObjectId();
            }
        }

        /// <summary>
        /// Récupère le nom du projet courant.
        /// </summary>
        /// <returns>Le nom du projet, ou <see cref="string.Empty"/> si non trouvé.</returns>
        public static string GetCurrentProjectName()
        {
            try
            {
                var projectPdm = GetCurrentProjectPdmObject();
                if (!projectPdm.IsEmpty)
                    return TSH.Pdm.GetName(projectPdm);
            }
            catch (Exception) { }

            return string.Empty;
        }

        #endregion

        #region Bibliothèques

        /// <summary>
        /// Récupère la liste des <see cref="PdmObjectId"/> de tous les projets bibliothèques PDM.
        /// </summary>
        /// <returns>
        /// Liste des <see cref="PdmObjectId"/> des bibliothèques, ou une liste vide si aucune n'est trouvée.
        /// </returns>
        /// <example>
        /// <code>
        /// foreach (var lib in PDM.GetLibraries())
        ///     Console.WriteLine(PDM.GetName(lib));
        /// </code>
        /// </example>
        public static List<PdmObjectId> GetLibraries()
        {
            var result = new List<PdmObjectId>();
            var allProjects = TSH.Pdm.GetProjects(false, true);
            if (allProjects == null) return result;

            foreach (var project in allProjects)
                if (!project.IsEmpty)
                    result.Add(project);

            return result;
        }

        /// <summary>
        /// Récupère les noms de tous les projets bibliothèques PDM.
        /// </summary>
        /// <returns>Liste des noms. Les projets dont le nom est illisible sont ignorés.</returns>
        /// <example>
        /// <code>
        /// foreach (var nom in PDM.GetLibraryNames())
        ///     Console.WriteLine(nom);
        /// </code>
        /// </example>
        public static List<string> GetLibraryNames()
        {
            var result = new List<string>();
            foreach (var lib in GetLibraries())
            {
                try { result.Add(TSH.Pdm.GetName(lib)); }
                catch { }
            }
            return result;
        }

        /// <summary>
        /// Recherche une bibliothèque PDM par son nom.
        /// </summary>
        /// <param name="name">Nom de la bibliothèque recherchée.</param>
        /// <param name="ignoreCase">Si <c>true</c>, la comparaison ignore la casse (par défaut : <c>true</c>).</param>
        /// <returns>
        /// Le <see cref="PdmObjectId"/> de la bibliothèque si trouvée ; sinon <see cref="PdmObjectId.Empty"/>.
        /// </returns>
        /// <example>
        /// <code>
        /// PdmObjectId lib = PDM.GetLibraryByName("docuType");
        /// if (!lib.IsEmpty)
        ///     Console.WriteLine("Bibliothèque trouvée !");
        /// </code>
        /// </example>
        public static PdmObjectId GetLibraryByName(string name, bool ignoreCase = true)
        {
            if (string.IsNullOrEmpty(name)) return PdmObjectId.Empty;

            var comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            foreach (var lib in GetLibraries())
            {
                try
                {
                    if (string.Equals(TSH.Pdm.GetName(lib), name, comparison))
                        return lib;
                }
                catch { }
            }
            return PdmObjectId.Empty;
        }

        /// <summary>
        /// Tente de récupérer une bibliothèque PDM par son nom.
        /// </summary>
        /// <param name="name">Nom recherché.</param>
        /// <param name="libraryId">Sortie : le <see cref="PdmObjectId"/> de la bibliothèque, ou <see cref="PdmObjectId.Empty"/> si non trouvée.</param>
        /// <param name="ignoreCase">Si <c>true</c>, la comparaison ignore la casse (par défaut : <c>true</c>).</param>
        /// <returns><c>true</c> si la bibliothèque a été trouvée ; sinon <c>false</c>.</returns>
        /// <example>
        /// <code>
        /// if (PDM.TryGetLibrary("docuType", out var libId))
        ///     Console.WriteLine(PDM.GetName(libId));
        /// </code>
        /// </example>
        public static bool TryGetLibrary(string name, out PdmObjectId libraryId, bool ignoreCase = true)
        {
            libraryId = GetLibraryByName(name, ignoreCase);
            return !libraryId.IsEmpty;
        }

        /// <summary>
        /// Ajoute des références à une ou plusieurs bibliothèques PDM pour le projet spécifié.
        /// </summary>
        /// <param name="ProjectId">Le <see cref="PdmObjectId"/> du projet cible. Si <see cref="PdmObjectId.IsEmpty"/> la méthode retourne immédiatement.</param>
        /// <param name="lib">Le <see cref="PdmObjectId"/> de la bibliothèque à référencer.</param>
        /// <param name="libs       
        /// Liste des <see cref="PdmObjectId"/> représentant les projets bibliothèques à référencer.
        /// Doit être non nulle ; si la liste est vide, aucun traitement n'est effectué.
        /// </param>
        /// <remarks>
        /// - La méthode capture et gère les exceptions en affichant une boîte de dialogue d'erreur (aucune exception n'est propagée).
        /// - Si <paramref name="ProjectId"/> est vide ou si <paramref name="lib"/> est vide, la méthode ne fait rien.
        /// - Cette méthode s'appuie sur l'API TopSolid via <c>TSH.Pdm.AddReferencedProjects</c>.
        /// </remarks>
        /// <example>
        /// <code>
        /// var projet = PDM.GetCurrentProjectPdmObject();
        /// var lib = PDM.GetLibraryByName("docuType");
        /// PDM.RefLibrary(projet, lib);
        /// </code>
        /// </example>
        public static void RefLibrary(PdmObjectId ProjectId, PdmObjectId lib)
        {
            if (ProjectId.IsEmpty) return;
            if (lib.IsEmpty) return;
            List<PdmObjectId> libs = new List<PdmObjectId> { lib };
            try
            {
                TSH.Pdm.AddReferencedProjects(ProjectId, libs);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERREUR lors du référencement de la bibliothèque dans le projet courant : {ex.Message}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Documents

        /// <summary>
        /// Récupère le <see cref="PdmObjectId"/> d'un document donné.
        /// </summary>
        /// <param name="documentId">L'identifiant du document.</param>
        /// <returns>Le <see cref="PdmObjectId"/> du document, ou <see cref="PdmObjectId.Empty"/> si échec.</returns>
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
        /// <returns><c>true</c> si le document est sous contrôle PDM ; sinon <c>false</c>.</returns>
        public static bool IsDocumentInPdm(DocumentId documentId)
            => !GetDocumentPdmObject(documentId).IsEmpty;

        /// <summary>
        /// Récupère tous les documents d'un projet PDM de manière récursive.
        /// </summary>
        /// <param name="projectPdm">Le <see cref="PdmObjectId"/> du projet à explorer.</param>
        /// <returns>Liste de tous les <see cref="PdmObjectId"/> de documents du projet.</returns>
        /// <example>
        /// <code>
        /// PdmObjectId lib = PDM.GetLibraryByName("docuType");
        /// var docs = PDM.GetAllProjectDocuments(lib);
        /// foreach (var doc in docs)
        ///     Console.WriteLine(PDM.GetName(doc));
        /// </code>
        /// </example>
        public static List<PdmObjectId> GetAllProjectDocuments(PdmObjectId projectPdm)
        {
            var allDocuments = new List<PdmObjectId>();
            try
            {
                if (!projectPdm.IsEmpty)
                {
                    TSH.Pdm.GetConstituents(projectPdm, out var dossiers, out var documents);
                    if (documents != null) allDocuments.AddRange(documents);
                    if (dossiers != null)
                        foreach (var dossier in dossiers)
                            allDocuments.AddRange(GetDocumentsRecursive(dossier));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERREUR : {ex.Message}\n\nStack:\n{ex.StackTrace}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return allDocuments;
        }

        /// <summary>
        /// Récupère tous les documents du projet courant de manière récursive.
        /// </summary>
        /// <returns>Liste de tous les <see cref="PdmObjectId"/> de documents du projet courant.</returns>
        public static List<PdmObjectId> GetAllProjectDocuments()
            => GetAllProjectDocuments(GetCurrentProjectPdmObject());

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
                if (documents != null) allDocuments.AddRange(documents);
                if (subFolders != null)
                    foreach (var subFolder in subFolders)
                        allDocuments.AddRange(GetDocumentsRecursive(subFolder));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ERREUR dans GetDocumentsRecursive : {ex.Message}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return allDocuments;
        }

        #endregion

        #region Dossiers

        /// <summary>
        /// Récupère la liste de tous les dossiers d'un projet PDM de manière récursive.
        /// </summary>
        /// <param name="projectPdm">Le <see cref="PdmObjectId"/> du projet à explorer.</param>
        /// <returns>Liste des <see cref="PdmObjectId"/> de tous les dossiers du projet.</returns>
        /// <example>
        /// <code>
        /// PdmObjectId lib = PDM.GetLibraryByName("docuType");
        /// var dossiers = PDM.GetProjectFolders(lib);
        /// foreach (var d in dossiers)
        ///     Console.WriteLine(PDM.GetName(d));
        /// </code>
        /// </example>
        public static List<PdmObjectId> GetProjectFolders(PdmObjectId projectPdm)
        {
            var folders = new List<PdmObjectId>();
            try
            {
                if (!projectPdm.IsEmpty)
                {
                    TSH.Pdm.GetConstituents(projectPdm, out var dossiers, out _);
                    if (dossiers != null)
                    {
                        folders.AddRange(dossiers);
                        foreach (var dossier in dossiers)
                            folders.AddRange(GetSubFoldersRecursive(dossier));
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
        /// Récupère la liste de tous les dossiers du projet courant de manière récursive.
        /// </summary>
        /// <returns>Liste des <see cref="PdmObjectId"/> de tous les dossiers du projet courant.</returns>
        public static List<PdmObjectId> GetProjectFolders()
            => GetProjectFolders(GetCurrentProjectPdmObject());

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
                TSH.Pdm.GetConstituents(folderPdm, out var subFolders, out _);
                if (subFolders != null)
                {
                    allSubFolders.AddRange(subFolders);
                    foreach (var subFolder in subFolders)
                        allSubFolders.AddRange(GetSubFoldersRecursive(subFolder));
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
        /// <param name="folderPdmObject">Le <see cref="PdmObjectId"/> du dossier.</param>
        /// <returns>Liste des <see cref="PdmObjectId"/> des fichiers du dossier uniquement.</returns>
        public static List<PdmObjectId> GetFilesInFolder(PdmObjectId folderPdmObject)
        {
            var files = new List<PdmObjectId>();
            try
            {
                if (!folderPdmObject.IsEmpty)
                {
                    TSH.Pdm.GetConstituents(folderPdmObject, out _, out var documents);
                    if (documents != null) files.AddRange(documents);
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
        /// Vérifie si un objet PDM est dans l'état coffré (<see cref="PdmObjectState.CheckedIn"/>).
        /// </summary>
        /// <param name="pdmObjectId">L'identifiant de l'objet PDM à vérifier.</param>
        /// <returns>
        /// <c>true</c> si l'objet est dans l'état <see cref="PdmObjectState.CheckedIn"/> ;
        /// sinon <c>false</c>.
        /// </returns>
        /// <example>
        /// <code>
        /// PdmObjectId pdmDoc = PDM.GetDocumentPdmObject(docId);
        /// if (!PDM.IsCheckedIn(pdmDoc))
        ///     // effectuer le coffrage
        /// </code>
        /// </example>
        public static bool IsCheckedIn(PdmObjectId pdmObjectId)
        {
            if (pdmObjectId.IsEmpty)
                return false;

            try
            {
                return TSH.Pdm.GetState(pdmObjectId) == PdmObjectState.CheckedIn;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Vérifie si un objet PDM est dans l'état validé (<see cref="PdmObjectState.Validated"/>).
        /// </summary>
        /// <param name="pdmObjectId">L'identifiant de l'objet PDM à vérifier.</param>
        /// <returns>
        /// <c>true</c> si l'objet est dans l'état <see cref="PdmObjectState.Validated"/> ;
        /// sinon <c>false</c>.
        /// </returns>
        /// <example>
        /// <code>
        /// PdmObjectId pdmDoc = PDM.GetDocumentPdmObject(docId);
        /// if (!PDM.IsValidated(pdmDoc))
        ///     // effectuer la validation
        /// </code>
        /// </example>
        public static bool IsValidated(PdmObjectId pdmObjectId)
        {
            if (pdmObjectId.IsEmpty)
                return false;

            try
            {
                return TSH.Pdm.GetState(pdmObjectId) == PdmObjectState.Validated;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Vérifie si un objet PDM est dans la corbeille (marqué pour suppression ou supprimé).
        /// </summary>
        /// <param name="pdmObjectId">L'identifiant de l'objet PDM à vérifier.</param>
        /// <returns>
        /// <c>true</c> si l'objet est dans l'état <see cref="PdmObjectState.ToDelete"/>,
        /// <see cref="PdmObjectState.Deleted"/> ou <see cref="PdmObjectState.LockedBecauseToDelete"/> ;
        /// sinon <c>false</c>.
        /// </returns>
        /// <example>
        /// <code>
        /// PdmObjectId pdmMaster = PDM.GetDocumentPdmObject(docMaster);
        /// if (PDM.IsInRecycleBin(pdmMaster))
        ///     Console.WriteLine("Le document est dans la corbeille.");
        /// </code>
        /// </example>
        public static bool IsInRecycleBin(PdmObjectId pdmObjectId)
        {
            if (pdmObjectId.IsEmpty)
                return false;

            try
            {
                PdmObjectState state = TSH.Pdm.GetState(pdmObjectId);
                return state == PdmObjectState.ToDelete
                    || state == PdmObjectState.Deleted
                    || state == PdmObjectState.LockedBecauseToDelete;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Récupère le nom d'un objet PDM.
        /// </summary>
        /// <param name="pdmObject">L'objet PDM.</param>
        /// <returns>Le nom de l'objet, ou <see cref="string.Empty"/> si échec.</returns>
        /// <example>
        /// <code>
        /// string nom = PDM.GetName(PDM.GetCurrentProjectPdmObject());
        /// </code>
        /// </example>
        public static string GetName(PdmObjectId pdmObject)
        {
            try
            {
                if (!pdmObject.IsEmpty)
                    return TSH.Pdm.GetName(pdmObject);
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
