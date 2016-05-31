using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace FIA_Biosum_Manager
{
    /// <summary>
    /// Interaction logic for XPSDocumentViewer.xaml
    /// </summary>
    public partial class XPSDocumentViewer : System.Windows.Navigation.NavigationWindow
    {
        public XPSDocumentViewer()
        {
            InitializeComponent();

        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Hide();
            e.Cancel = true;
        }
    }
}
