using System;
using TopSolid.Kernel.Automating;
using TopSolid.Cad.Design.Automating;
using TopSolid.Cam.NC.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;
using TSHD = TopSolid.Cad.Design.Automating.TopSolidDesignHost;
using TSCH = TopSolid.Cam.NC.Kernel.Automating.TopSolidCamHost;
using Wpf = System.Windows;

namespace TSCamPgmRename_WPF
{
    public class StartConnect
    {
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

        public void ConnectionTopsolid()
        {
            ConnectToTopSolid();
            ConnectToTopSolidDesignHost();
            ConnectToTopSolidCamHost();
        }
    }
}
