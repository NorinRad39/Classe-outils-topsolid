using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TopSolid.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;

namespace OutilsTs
{
    /// <summary>
    /// Classe utilitaire pour la gestion des opérations PDM (Product Data Management) dans TopSolid.
    /// </summary>
    public static class PDM
    {
        #region Récupération Projet

        /// <summary>
        /// Récupère le PdmObjectId du projet courant.
        /// Si un document est ouvert, utilise son projet parent.
        /// Sinon, utilise le projet actif dans TopSolid.
        /// </summary>
        /// <returns>Le PdmObjectId du projet courant, ou un PdmObjectId vide si aucun projet n'est trouvé.</returns>
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
        public static bool IsDocumentInPdm(DocumentId documentId)
        {
            var pdmObject = GetDocumentPdmObject(documentId);
            return !pdmObject.IsEmpty;
        }

        #endregion

        #region Gestion Dossiers (à implémenter)

        /// <summary>
        /// Récupère la liste des dossiers du projet courant.
        /// </summary>
        /// <returns>Liste des PdmObjectId des dossiers.</returns>
        public static List<PdmObjectId> GetProjectFolders()
        {
            var folders = new List<PdmObjectId>();

            try
            {
                var projectPdm = GetCurrentProjectPdmObject();
                if (!projectPdm.IsEmpty)
                {
                    // TODO: Implémenter la récupération des dossiers
                    // folders = TSH.Pdm.GetChildren(projectPdm)
                    //               .Where(child => TSH.Pdm.IsFolder(child))
                    //               .ToList();
                }
            }
            catch (Exception)
            {
                // Retourne liste vide en cas d'erreur
            }

            return folders;
        }

        /// <summary>
        /// Récupère la liste des fichiers dans un dossier PDM donné.
        /// </summary>
        /// <param name="folderPdmObject">Le PdmObjectId du dossier.</param>
        /// <returns>Liste des PdmObjectId des fichiers.</returns>
        public static List<PdmObjectId> GetFilesInFolder(PdmObjectId folderPdmObject)
        {
            var files = new List<PdmObjectId>();

            try
            {
                if (!folderPdmObject.IsEmpty)
                {
                    // TODO: Implémenter la récupération des fichiers
                    // files = TSH.Pdm.GetChildren(folderPdmObject)
                    //             .Where(child => !TSH.Pdm.IsFolder(child))
                    //             .ToList();
                }
            }
            catch (Exception)
            {
                // Retourne liste vide en cas d'erreur
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
        public static string GetName(PdmObjectId pdmObject)
        {
            try
            {
                if (!pdmObject.IsEmpty)
                {
                    return TSH.Pdm.GetName(pdmObject);
                }
            }
            catch (Exception)
            {
                // Silencieux
            }

            return string.Empty;
        }

        #endregion
    }
}
