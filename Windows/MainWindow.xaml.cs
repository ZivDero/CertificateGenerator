using System.Windows;
using CertificateGenerator.ViewModel;

namespace CertificateGenerator.Windows
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            ((MainViewModel) DataContext).InitializeViewModel();
        }
    }
}
