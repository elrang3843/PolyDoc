using System.Windows;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

/// <summary>표의 특정 행 속성을 편집하는 다이얼로그 (높이, 머리글 여부).</summary>
public partial class RowPropertiesWindow : Window
{
    private readonly Table _table;
    private readonly int _rowIndex;

    public RowPropertiesWindow(Table table, int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= table.Rows.Count)
            throw new ArgumentOutOfRangeException(nameof(rowIndex));

        InitializeComponent();
        _table = table;
        _rowIndex = rowIndex;
        LoadValues();
        WireUpEvents();
    }

    private void LoadValues()
    {
        var row = _table.Rows[_rowIndex];
        HeightBox.Text = row.HeightMm > 0 ? row.HeightMm.ToString("F1") : "0";
        IsHeaderCheck.IsChecked = row.IsHeader;
    }

    private void WireUpEvents()
    {
        OkButton.Click += OkButton_Click;
        CancelButton.Click += CancelButton_Click;
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        if (!double.TryParse(HeightBox.Text, out double height))
        {
            MessageBox.Show("높이는 숫자여야 합니다.", "입력 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var row = _table.Rows[_rowIndex];
        row.HeightMm = height;
        row.IsHeader = IsHeaderCheck.IsChecked ?? false;

        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
