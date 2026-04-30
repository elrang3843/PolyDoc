using System.Windows;
using PolyDonky.Iwpf;
using PolyDonky.App.Models;
using SR = PolyDonky.App.Properties.Resources;

namespace PolyDonky.App.Views;

public partial class DocumentInfoWindow : Window
{
    private readonly DocumentInfoModel _model;

    public DocumentInfoWindow(DocumentInfoModel model)
    {
        _model = model;
        InitializeComponent();
        DataContext = model;
        UpdatePasswordStatus();
        UpdateAuthorEditability();
    }

    private void UpdateAuthorEditability()
    {
        if (_model.HasBeenSaved && !string.IsNullOrWhiteSpace(_model.Author))
            AuthorField.IsEnabled = false;
    }

    private void UpdatePasswordStatus()
        => PasswordStatusText.Text = _model.PasswordMode switch
        {
            PasswordMode.Read  => SR.DocInfoPwdModeRead,
            PasswordMode.Write => SR.DocInfoPwdModeWrite,
            PasswordMode.Both  => SR.DocInfoPwdModeBoth,
            _                  => SR.DocInfoPwdModeNone,
        };

    private void OnPasswordChange(object sender, RoutedEventArgs e)
    {
        var dlg = new PasswordChangeWindow(_model.PasswordMode) { Owner = this };
        if (dlg.ShowDialog() != true) return;

        _model.PasswordChanged     = true;
        _model.PasswordMode        = dlg.ResultMode;
        _model.NewReadPassword     = dlg.ResultReadPassword;
        _model.NewWritePassword    = dlg.ResultWritePassword;
        UpdatePasswordStatus();
    }

    private void OnOk(object sender, RoutedEventArgs e)     => DialogResult = true;
    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}
