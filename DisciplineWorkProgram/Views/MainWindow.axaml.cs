using Avalonia.Controls;
using DisciplineWorkProgram.ViewModels;

namespace DisciplineWorkProgram.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainWindowViewModel();
        }
    }
}