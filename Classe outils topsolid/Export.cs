using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TopSolid.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;
using TSHD = TopSolid.Cad.Design.Automating.TopSolidDesignHost;

/*
================================================================================
Liste des méthodes disponibles dans la classe Export (OutilsTs)
================================================================================

Méthodes publiques :
- void ExportDocId(DocumentId documentId, string cheminDossier, string nomFichier, string extension, Dictionary<string, string> customOptions = null)
    // Exporte un document TopSolid au format déterminé par l'extension.
    // Syntaxe :
    // Export.ExportDocId(docId, @"C:\Export", "piece", "x_t", new Dictionary<string, string> { {"SAVE_VERSION", "30"} });
    // Export.ExportDocId(docId, @"C:\Export", "piece", "step");

Méthodes privées :
- bool FindExporterIndexByExtension(string extension, out int exporterIndex)
    // Trouve l'exporteur correspondant à une extension de fichier.
- List<KeyValue> ApplyCustomOptions(int exporterIndex, Dictionary<string, string> customOptions)
    // Applique des options personnalisées à l'exporteur.

================================================================================
*/

namespace OutilsTs
{
    /// <summary>
    /// Classe utilitaire pour gérer les exports TopSolid.
    /// Détecte automatiquement l'exporteur en fonction de l'extension du fichier.
    /// </summary>
    public static class Export
    {
        #region Méthodes publiques

        /// <summary>
        /// Exporte un document TopSolid au format déterminé par l'extension.
        /// </summary>
        /// <param name="documentId">Identifiant du document à exporter.</param>
        /// <param name="cheminDossier">Chemin du dossier de destination (ex: @"C:\Export").</param>
        /// <param name="nomFichier">Nom du fichier sans extension (ex: "piece").</param>
        /// <param name="extension">Extension du fichier (ex: "x_t", "step", "igs").</param>
        /// <param name="customOptions">Dictionnaire d'options personnalisées (ex: {"SAVE_VERSION", "30"}). Si null, utilise les options par défaut.</param>
        /// <remarks>
        /// Namespace: OutilsTs
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// // Export Parasolid X_T version 30
        /// var options = new Dictionary&lt;string, string&gt; { {"SAVE_VERSION", "30"} };
        /// Export.ExportDocId(docId, @"C:\Export", "piece", "x_t", options);
        /// 
        /// // Export STEP sans options
        /// Export.ExportDocId(docId, @"C:\Export", "piece", "step");
        /// </code>
        /// </example>
        public static void ExportDocId(DocumentId documentId, string cheminDossier, string nomFichier, string extension, Dictionary<string, string> customOptions = null)
        {
            // 1. Vérifier les paramètres
            if (string.IsNullOrEmpty(cheminDossier))
            {
                throw new ArgumentException("Le chemin du dossier ne peut pas être vide.", nameof(cheminDossier));
            }

            if (string.IsNullOrEmpty(nomFichier))
            {
                throw new ArgumentException("Le nom du fichier ne peut pas être vide.", nameof(nomFichier));
            }

            if (string.IsNullOrEmpty(extension))
            {
                throw new ArgumentException("L'extension ne peut pas être vide.", nameof(extension));
            }

            // 2. Nettoyer l'extension (enlever le point si présent)
            extension = extension.TrimStart('.').ToLower();

            // 3. Construire le chemin complet
            string cheminComplet = Path.Combine(cheminDossier, $"{nomFichier}.{extension}");

            // 4. Trouver l'exporteur correspondant à l'extension
            if (!FindExporterIndexByExtension(extension, out int exporterIndex))
            {
                throw new InvalidOperationException($"Aucun exporteur trouvé pour l'extension '.{extension}'.");
            }

            // 5. Effectuer l'export (avec ou sans options)
            if (customOptions != null && customOptions.Count > 0)
            {
                // Export avec options personnalisées
                List<KeyValue> options = ApplyCustomOptions(exporterIndex, customOptions);
                TSH.Documents.ExportWithOptions(exporterIndex, options, documentId, cheminComplet);
            }
            else
            {
                // Export simple (sans options personnalisées)
                TSH.Documents.Export(exporterIndex, documentId, cheminComplet);
            }
        }

        /// <summary>
        /// Exporte le document courant TopSolid.
        /// </summary>
        /// <param name="cheminDossier">Chemin du dossier de destination (ex: @"C:\Export").</param>
        /// <param name="nomFichier">Nom du fichier sans extension (ex: "piece").</param>
        /// <param name="extension">Extension du fichier (ex: "x_t", "step", "igs").</param>
        /// <param name="customOptions">Dictionnaire d'options personnalisées (ex: {"SAVE_VERSION", "30"}). Si null, utilise les options par défaut.</param>
        /// <remarks>
        /// Namespace: OutilsTs
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// // Exporter le document courant en STEP
        /// Export.ExportCurrent(@"C:\Export", "document_courant", "step");
        /// </code>
        /// </example>
        public static void ExportCurrent(string cheminDossier, string nomFichier, string extension, Dictionary<string, string> customOptions = null)
        {
            DocumentId currentDoc = TSH.Documents.EditedDocument;
            ExportDocId(currentDoc, cheminDossier, nomFichier, extension, customOptions);
        }

        #endregion

        #region Méthodes privées

        /// <summary>
        /// Recherche l'index de l'exporteur correspondant à une extension de fichier.
        /// </summary>
        /// <param name="extension">Extension du fichier (sans le point, ex: "x_t", "step").</param>
        /// <param name="exporterIndex">Index de l'exporteur trouvé (ou -1 si non trouvé).</param>
        /// <returns>True si un exporteur a été trouvé, False sinon.</returns>
        private static bool FindExporterIndexByExtension(string extension, out int exporterIndex)
        {
            exporterIndex = -1;

            for (int i = 0; i < TSH.Application.ExporterCount; i++)
            {
                TSH.Application.GetExporterFileType(i, out string fileTypeName, out string[] outFileExtensions);
                
                // Vérifier si l'extension existe dans le tableau
                if (outFileExtensions != null && outFileExtensions.Any(ext => ext.TrimStart('.').Equals(extension, StringComparison.OrdinalIgnoreCase)))
                {
                    exporterIndex = i;
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Applique des options personnalisées aux options par défaut de l'exporteur.
        /// Ne modifie que les clés spécifiées dans customOptions.
        /// </summary>
        /// <param name="exporterIndex">Index de l'exporteur.</param>
        /// <param name="customOptions">Dictionnaire des options à personnaliser (clé-valeur).</param>
        /// <returns>Liste des options configurées pour l'export.</returns>
        private static List<KeyValue> ApplyCustomOptions(int exporterIndex, Dictionary<string, string> customOptions)
        {
            List<KeyValue> options = TSH.Application.GetExporterOptions(exporterIndex);

            // Parcourir chaque option personnalisée
            foreach (var customOption in customOptions)
            {
                bool optionFound = false;

                // Chercher et modifier l'option si elle existe
                for (int i = 0; i < options.Count; i++)
                {
                    if (options[i].Key == customOption.Key)
                    {
                        options[i] = new KeyValue(customOption.Key, customOption.Value);
                        optionFound = true;
                        break;
                    }
                }

                // Avertissement si l'option n'existe pas pour cet exporteur
                if (!optionFound)
                {
                    System.Diagnostics.Debug.WriteLine($"Avertissement : L'option '{customOption.Key}' n'existe pas pour cet exporteur.");
                }
            }

            return options;
        }

        #endregion
    }
}