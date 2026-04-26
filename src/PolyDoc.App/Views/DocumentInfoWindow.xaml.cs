using System.Windows;
using PolyDoc.App.Models;

namespace PolyDoc.App.Views;

public partial class DocumentInfoWindow : Window
{
    public DocumentInfoWindow(DocumentInfoModel model)
    {
        InitializeComponent();
        DataContext = model;
    }

    private void OnClose(object sender, RoutedEventArgs e) => Close();
}
