using System;
using Wpf = System.Windows;
using WinForms = System.Windows.Forms;

/*
================================================================================
Liste des méthodes et types disponibles dans la classe UiHelper (OutilsTs)
================================================================================

Méthodes publiques :
- static void ShowMessage(string text, string title, MessageType type = MessageType.Info)
    // Affiche une boîte de dialogue (WPF ou WinForms) avec le texte, le titre et le type d’icône.
    // Syntaxe :
    // UiHelper.ShowMessage("Message", "Titre", MessageType.Warning);

Types internes :
- enum MessageType
    // Type de message à afficher (Info, Warning, Error).

Méthodes privées :
- static bool IsWpfProject()
    // Détecte si l’application est WPF.

================================================================================
Pour toute nouvelle méthode ou type, ajoutez-la à cette liste pour faciliter la découverte.
================================================================================
*/

namespace OutilsTs
{
    /// <summary>
    /// Fournit des méthodes utilitaires pour afficher des messages à l'utilisateur
    /// dans des applications WPF ou WinForms.
    /// </summary>
    public static class UiHelper
    {
        #region Méthodes publiques

        /// <summary>
        /// Affiche une boîte de dialogue à l'utilisateur avec le texte, le titre et le type d'icône spécifiés.
        /// Utilise automatiquement WPF ou WinForms selon le contexte d'exécution.
        /// </summary>
        /// <param name="text">Texte du message à afficher.</param>
        /// <param name="title">Titre de la boîte de dialogue.</param>
        /// <param name="type">Type d'icône à afficher (Info, Warning, Error). Par défaut : Info.</param>
        /// <remarks>
        /// Namespace: OutilsTs<br/>
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// UiHelper.ShowMessage("Opération terminée", "Information", MessageType.Info);
        /// UiHelper.ShowMessage("Attention !", "Avertissement", MessageType.Warning);
        /// UiHelper.ShowMessage("Erreur critique", "Erreur", MessageType.Error);
        /// </code>
        /// </example>
        public static void ShowMessage(string text, string title,
            MessageType type = MessageType.Info)
        {
            if (IsWpfProject())
            {
                Wpf.MessageBoxImage icon = type switch
                {
                    MessageType.Error => Wpf.MessageBoxImage.Error,
                    MessageType.Warning => Wpf.MessageBoxImage.Warning,
                    _ => Wpf.MessageBoxImage.Information
                };

                Wpf.MessageBox.Show(text, title, Wpf.MessageBoxButton.OK, icon);
            }
            else
            {
                WinForms.MessageBoxIcon icon = type switch
                {
                    MessageType.Error => WinForms.MessageBoxIcon.Error,
                    MessageType.Warning => WinForms.MessageBoxIcon.Warning,
                    _ => WinForms.MessageBoxIcon.Information
                };

                WinForms.MessageBox.Show(text, title, WinForms.MessageBoxButtons.OK, icon);
            }
        }

        #endregion

        #region Méthodes privées

        /// <summary>
        /// Détecte si l'application courante est une application WPF.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs<br/>
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// bool isWpf = UiHelper.IsWpfProject();
        /// </code>
        /// </example>
        /// <returns>
        /// Type: <see cref="bool"/>
        /// <c>true</c> si l'application est WPF, sinon <c>false</c>.
        /// </returns>
        private static bool IsWpfProject()
        {
            try
            {
                var appType = Type.GetType("System.Windows.Application, PresentationFramework");
                return appType != null;
            }
            catch
            {
                return false;
            }
        }

        #endregion
    }

    /// <summary>
    /// Spécifie le type de message à afficher dans la boîte de dialogue utilisateur.
    /// </summary>
    /// <remarks>
    /// Namespace: OutilsTs<br/>
    /// Assembly: OutilsTs (in OutilsTs.dll)
    /// </remarks>
    public enum MessageType
    {
        /// <summary>
        /// Message d'information.
        /// </summary>
        Info,
        /// <summary>
        /// Message d'avertissement.
        /// </summary>
        Warning,
        /// <summary>
        /// Message d'erreur.
        /// </summary>
        Error
    }
}
