using System.Windows;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

/// <summary>
/// 표 삽입 다이얼로그. 행·열 수와 너비(mm)를 입력받아 <see cref="ResultTable"/> 을 생성한다.
/// </summary>
public partial class TableInsertDialog : Window
{
    public Table? ResultTable { get; private set; }

    public TableInsertDialog()
    {
        InitializeComponent();
        Loaded += (_, _) => RowsBox.Focus();
    }

    private void OnInsertClick(object sender, RoutedEventArgs e)
    {
        if (!int.TryParse(RowsBox.Text.Trim(), out int rows) || rows < 1 || rows > 100)
        {
            MessageBox.Show(this, "행 수를 1 ~ 100 사이로 입력하세요.", "표 삽입",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            RowsBox.Focus();
            return;
        }
        if (!int.TryParse(ColsBox.Text.Trim(), out int cols) || cols < 1 || cols > 20)
        {
            MessageBox.Show(this, "열 수를 1 ~ 20 사이로 입력하세요.", "표 삽입",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            ColsBox.Focus();
            return;
        }
        if (!double.TryParse(WidthBox.Text.Trim(), out double widthMm) || widthMm < 0)
        {
            MessageBox.Show(this, "너비를 0 이상의 숫자로 입력하세요 (0 = 자동).", "표 삽입",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            WidthBox.Focus();
            return;
        }

        bool headerRow = HeaderRowCheck.IsChecked == true;
        ResultTable = BuildTable(rows, cols, widthMm, headerRow);
        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e) => Close();

    private static Table BuildTable(int rows, int cols, double totalWidthMm, bool headerRow)
    {
        double colWidth = totalWidthMm > 0 ? totalWidthMm / cols : 0;
        var table = new Table
        {
            BorderThicknessPt = 0.75,      // 기본 테두리 두께
            BorderColor = "#C8C8C8",       // 기본 연회색
        };

        for (int c = 0; c < cols; c++)
            table.Columns.Add(new TableColumn { WidthMm = colWidth });

        for (int r = 0; r < rows; r++)
        {
            var row = new TableRow { IsHeader = headerRow && r == 0 };
            for (int c = 0; c < cols; c++)
                row.Cells.Add(new TableCell { Blocks = { Paragraph.Of(string.Empty) } });
            table.Rows.Add(row);
        }

        return table;
    }
}
