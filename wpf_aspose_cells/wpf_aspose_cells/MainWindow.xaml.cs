using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace wpf_aspose_cells
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            #region build doc and sheet1
            var doc = new Aspose.Cells.Workbook();
            doc.Worksheets.Clear(); //start with 3 default sheets
            var sheet1 = doc.Worksheets.Add("sheet1");
            #endregion

            #region //1.set the border for single cell style
            if (false)
            {
                var cell = sheet1.Cells[1, 1];
                var style = cell.GetStyle();
                cell.PutValue("szp");
                style.SetBorder(Aspose.Cells.BorderType.BottomBorder, Aspose.Cells.CellBorderType.Medium, System.Drawing.Color.Red);
                style.SetBorder(Aspose.Cells.BorderType.TopBorder, Aspose.Cells.CellBorderType.Medium, System.Drawing.Color.Red);
                style.SetBorder(Aspose.Cells.BorderType.LeftBorder, Aspose.Cells.CellBorderType.Medium, System.Drawing.Color.Red);
                style.SetBorder(Aspose.Cells.BorderType.RightBorder, Aspose.Cells.CellBorderType.Medium, System.Drawing.Color.Red);
                cell.SetStyle(style);
            }
            #endregion

            #region //2.set the border for multi cells style (batch set)
            if (false)
            {
                var range = sheet1.Cells.CreateRange(2, 2, 3, 3);
                var style = sheet1.Cells[2, 2].GetStyle();
                range.PutValue("test1", true, true);
                style.SetBorder(Aspose.Cells.BorderType.BottomBorder, Aspose.Cells.CellBorderType.Medium, System.Drawing.Color.Red);
                style.SetBorder(Aspose.Cells.BorderType.TopBorder, Aspose.Cells.CellBorderType.Medium, System.Drawing.Color.Red);
                style.SetBorder(Aspose.Cells.BorderType.LeftBorder, Aspose.Cells.CellBorderType.Medium, System.Drawing.Color.Red);
                style.SetBorder(Aspose.Cells.BorderType.RightBorder, Aspose.Cells.CellBorderType.Medium, System.Drawing.Color.Red);
                //range.SetStyle(style);
                range.ApplyStyle(style, new Aspose.Cells.StyleFlag() { Borders = true, BottomBorder = true });
            }
            #endregion

            #region //3.set the border around the range.
            if (false)
            {
                var range = sheet1.Cells.CreateRange(2, 2, 3, 3);
                range.PutValue("1", true, true);
                var border_type = Aspose.Cells.CellBorderType.Dotted;
                range.SetOutlineBorder(Aspose.Cells.BorderType.TopBorder, border_type, System.Drawing.Color.Red);
                range.SetOutlineBorder(Aspose.Cells.BorderType.BottomBorder, border_type, System.Drawing.Color.Red);
                range.SetOutlineBorder(Aspose.Cells.BorderType.LeftBorder, border_type, System.Drawing.Color.Red);
                range.SetOutlineBorder(Aspose.Cells.BorderType.RightBorder, border_type, System.Drawing.Color.Red);
            }
            #endregion

            #region //4.import from datatable.(sheet1.AutoFitColumns())
            if (true)
            {
                var dt = new DataTable();
                var c1 = dt.Columns.Add("ID");
                var c2 = dt.Columns.Add("Name");
                var c3 = dt.Columns.Add("Address");

                dt.Rows.Add("1", "szp", "china first street1");
                dt.Rows.Add("2", "szp", "great wall");
                dt.Rows.Add("3", "szp", "zgc\\ great wall");
                dt.Rows.Add("3", "szp", "dlxc");

                sheet1.Cells.ImportData(dt, 0, 0, new Aspose.Cells.ImportTableOptions() { IsFieldNameShown = true, ConvertNumericData = true, CheckMergedCells = true });
                sheet1.IsGridlinesVisible = false;
                sheet1.AutoFitColumns();
                sheet1.AutoFitRows();
            }

            #endregion

            #region Save to file
            var path = @"d:\1.xlsx";
            try
            {
                doc.Save(path, Aspose.Cells.SaveFormat.Xlsx);
                this.Title = String.Format("Export to {0} successfully!", path);
            }
            catch
            {
                this.Title = String.Format("faild to Export to {0}!", path);
            }
            #endregion
        }
        private Aspose.Cells.Cell GetDown(Aspose.Cells.Worksheet sheet, Aspose.Cells.Cell cell)
        {
            var row = cell.Row;
            var col = cell.Column;
            Aspose.Cells.Cell c = null;
            while ((c = sheet.Cells[++row, col]) == null) ;
            return c;
        }
    }
}
