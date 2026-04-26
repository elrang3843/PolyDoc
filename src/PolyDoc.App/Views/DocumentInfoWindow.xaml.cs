using System.Windows;
using PolyDoc.Iwpf;
using PolyDoc.App.Models;
using SR = PolyDoc.App.Properties.Resources;

namespace PolyDoc.App.Views;

public partial class DocumentInfoWindow : Window
{
    private readonly DocumentInfoModel _model;

    public DocumentInfoWindow(DocumentInfoModel model)
    {
        _model = model;
        InitializeComponent();
        DataContext = model;
        UpdatePasswordStatus();
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
