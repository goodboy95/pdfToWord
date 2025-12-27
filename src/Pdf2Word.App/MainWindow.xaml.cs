using System.Windows;
using Pdf2Word.App.ViewModels;

namespace Pdf2Word.App;

public partial class MainWindow : Window
{
    public MainWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
    }
}
