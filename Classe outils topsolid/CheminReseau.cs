using System;
using System.Runtime.InteropServices;
using System.Text;

/*
================================================================================
Liste des méthodes disponibles dans la classe CheminReseau (OutilsTs)
================================================================================

Méthodes publiques :
- string ConvertToUncPath(string path)
    // Convertit un chemin de lecteur mappé (ex: Z:\) en chemin UNC (ex: \\serveur\partage\).
    // Syntaxe :
    // string uncPath = CheminReseau.ConvertToUncPath(@"Z:\dossier\fichier.txt");

Méthodes privées :
- int WNetGetConnection(string localName, StringBuilder remoteName, ref int length)
    // API Windows pour obtenir le chemin UNC d'un lecteur réseau mappé.

================================================================================
*/

namespace OutilsTs
{
    /// <summary>
    /// Classe utilitaire pour la gestion des chemins réseau et conversions UNC.
    /// </summary>
    public static class CheminReseau
    {
        #region Méthodes privées (P/Invoke)

        /// <summary>
        /// API Windows pour récupérer le chemin UNC d'un lecteur réseau mappé.
        /// </summary>
        /// <param name="localName">Lettre du lecteur (ex: "Z:").</param>
        /// <param name="remoteName">Buffer pour recevoir le chemin UNC.</param>
        /// <param name="length">Taille du buffer.</param>
        /// <returns>0 si succès, code d'erreur sinon.</returns>
        [DllImport("mpr.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int WNetGetConnection(
            [MarshalAs(UnmanagedType.LPWStr)] string localName,
            [MarshalAs(UnmanagedType.LPWStr)] StringBuilder remoteName,
            ref int length);

        #endregion

        #region Méthodes publiques

        /// <summary>
        /// Convertit un chemin de lecteur mappé en chemin UNC.
        /// Si le chemin n'est pas un lecteur réseau, retourne le chemin original.
        /// </summary>
        /// <param name="path">Chemin à convertir (ex: @"Z:\dossier\fichier.txt").</param>
        /// <returns>Chemin UNC (ex: @"\\serveur\partage\dossier\fichier.txt") ou le chemin original.</returns>
        /// <remarks>
        /// Namespace: OutilsTs
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// // Lecteur mappé Z: → \\serveur\partage\
        /// string uncPath = CheminReseau.ConvertToUncPath(@"Z:\projet\piece.TopSolid_Part");
        /// // Résultat: \\serveur\partage\projet\piece.TopSolid_Part
        /// 
        /// // Chemin local (non mappé)
        /// string localPath = CheminReseau.ConvertToUncPath(@"C:\Users\Documents\fichier.txt");
        /// // Résultat: C:\Users\Documents\fichier.txt (inchangé)
        /// </code>
        /// </example>
        public static string ConvertToUncPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.Length < 2)
                return path;

            // Vérifier si c'est un chemin avec lettre de lecteur (ex: Z:\)
            if (path[1] == ':')
            {
                string driveLetter = path.Substring(0, 2);
                StringBuilder uncPath = new StringBuilder(512);
                int size = uncPath.Capacity;

                int result = WNetGetConnection(driveLetter, uncPath, ref size);

                if (result == 0)
                {
                    // C'est un lecteur réseau mappé, remplacer par le chemin UNC
                    return uncPath.ToString() + path.Substring(2);
                }
            }

            // Retourner le chemin original si ce n'est pas un lecteur réseau
            return path;
        }

        #endregion
    }
}