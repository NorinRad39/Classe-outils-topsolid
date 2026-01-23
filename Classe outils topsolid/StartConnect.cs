using System;
using TopSolid.Kernel.Automating;
using TopSolid.Cad.Design.Automating;
using TopSolid.Cam.NC.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;
using TSHD = TopSolid.Cad.Design.Automating.TopSolidDesignHost;
using TSCH = TopSolid.Cam.NC.Kernel.Automating.TopSolidCamHost;
using Wpf = System.Windows;

/*
================================================================================
Liste des méthodes disponibles dans la classe StartConnect (OutilsTs)
================================================================================

Méthodes publiques :
- void ConnectionTopsolid()
    // Connecte aux hôtes TopSolid Kernel, Design et CAM.
    // Syntaxe :
    // var connector = new StartConnect();
    // connector.ConnectionTopsolid();

Méthodes privées :
- void ConnectToTopSolid()
    // Connecte à TopSolid Kernel.
- void ConnectToTopSolidDesignHost()
    // Connecte à TopSolid Design.
- void ConnectToTopSolidCamHost()
    // Connecte à TopSolid CAM.

================================================================================
Pour toute nouvelle méthode, ajoutez-la à cette liste pour faciliter la découverte.
================================================================================
*/

namespace OutilsTs
{
    /// <summary>
    /// Classe utilitaire pour gérer la connexion aux différents hôtes TopSolid (Kernel, Design, CAM).
    /// </summary>
    public class StartConnect
    {
        #region Méthodes privées

        /// <summary>
        /// Connecte à TopSolid Kernel.
        /// Affiche un message en cas d'échec ou si déjà connecté.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// // Appel interne depuis ConnectionTopsolid()
        /// ConnectToTopSolid();
        /// </code>
        /// </example>
        private void ConnectToTopSolid()
        {
            try
            {
                if (!TSH.IsConnected)
                {
                    TSH.Connect();

                    if (!TSH.IsConnected)
                        Wpf.MessageBox.Show("Connexion échouée à TopSolid Kernel.", "Erreur", Wpf.MessageBoxButton.OK, Wpf.MessageBoxImage.Error);
                }
                else
                {
                    Wpf.MessageBox.Show("TopSolid Kernel est déjà connecté.", "Information", Wpf.MessageBoxButton.OK, Wpf.MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                Wpf.MessageBox.Show($"Erreur TopSolid Kernel : {ex.Message}", "Erreur", Wpf.MessageBoxButton.OK, Wpf.MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Connecte à TopSolid Design.
        /// Affiche un message en cas d'échec ou si déjà connecté.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// // Appel interne depuis ConnectionTopsolid()
        /// ConnectToTopSolidDesignHost();
        /// </code>
        /// </example>
        private void ConnectToTopSolidDesignHost()
        {
            try
            {
                if (!TSHD.IsConnected)
                {
                    TSHD.Connect();

                    if (!TSHD.IsConnected)
                        Wpf.MessageBox.Show("Connexion échouée à TopSolid Design.", "Erreur", Wpf.MessageBoxButton.OK, Wpf.MessageBoxImage.Error);
                }
                else
                {
                    Wpf.MessageBox.Show("TopSolid Design est déjà connecté.", "Information", Wpf.MessageBoxButton.OK, Wpf.MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                Wpf.MessageBox.Show($"Erreur TopSolid Design : {ex.Message}", "Erreur", Wpf.MessageBoxButton.OK, Wpf.MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Connecte à TopSolid CAM.
        /// Affiche un message en cas d'échec ou si déjà connecté.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// // Appel interne depuis ConnectionTopsolid()
        /// ConnectToTopSolidCamHost();
        /// </code>
        /// </example>
        private void ConnectToTopSolidCamHost()
        {
            try
            {
                if (!TSCH.IsConnected)
                {
                    TSCH.Connect();

                    if (!TSCH.IsConnected)
                        Wpf.MessageBox.Show("Connexion échouée à TopSolid CAM.", "Erreur", Wpf.MessageBoxButton.OK, Wpf.MessageBoxImage.Error);
                }
                else
                {
                    Wpf.MessageBox.Show("TopSolid CAM est déjà connecté.", "Information", Wpf.MessageBoxButton.OK, Wpf.MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                Wpf.MessageBox.Show($"Erreur TopSolid CAM : {ex.Message}", "Erreur", Wpf.MessageBoxButton.OK, Wpf.MessageBoxImage.Error);
            }
        }

        #endregion

        #region Méthodes publiques

        /// <summary>
        /// Connecte aux hôtes TopSolid Kernel, Design et CAM.
        /// </summary>
        /// <remarks>
        /// Namespace: OutilsTs
        /// Assembly: OutilsTs (in OutilsTs.dll)
        /// </remarks>
        /// <example>
        /// <code>
        /// var connector = new StartConnect();
        /// connector.ConnectionTopsolid();
        /// </code>
        /// </example>
        public void ConnectionTopsolid()
        {
            ConnectToTopSolid();
            ConnectToTopSolidDesignHost();
            ConnectToTopSolidCamHost();
        }

        #endregion
    }
}
