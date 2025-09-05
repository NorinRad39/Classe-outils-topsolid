using System;
using Wpf = System.Windows;
using WinForms = System.Windows.Forms;

namespace OutilsTs
{
    public static class UiHelper
    {
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

        // Cette fonction peut être simple : on considère que si System.Windows.Application existe, c'est WPF
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
    }

    public enum MessageType
    {
        Info,
        Warning,
        Error
    }
}
