using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
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

namespace wpf_aspose_pdf
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
            this.CreateTableOnPdfFile_4();
        }
        private void CreateTableOnPdfFile_4()
        {
            //create filename of pdf_file_tosave 
            var pdf_file_tosave = @"d:\1.pdf";

            //new pdf doc
            var pdfDoc = new Aspose.Pdf.Document();

            //editing
            pdfDoc.SetTitle("hello!");
            var pdfPage = pdfDoc.Pages.Add();

            var table = new Aspose.Pdf.Table();
            // Set the table border color as LightGray
            table.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, .5f, Aspose.Pdf.Color.LightGray);
            // Set the border for table cells
            table.DefaultCellBorder = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, .5f, Aspose.Pdf.Color.LightGray);
            // Create a loop to add 10 rows
            for (int row_count = 1; row_count <= 10; row_count++)
            {
                // Add row to table
                Aspose.Pdf.Row row = table.Rows.Add();
                // Add table cells
                row.Cells.Add("Column (" + row_count + ", 1)");
                row.Cells.Add("Column (" + row_count + ", 2)");
            }
            // Add table object to first page of input document
            pdfPage.Paragraphs.Add(table);
            //Aspose.Pdf.Layer layer1 = new Aspose.Pdf.Layer("layer1","layer1");
            //pdfPage.Layers.Add(layer1); //why is pdfPage.layers null?

            //save
            try
            {
                pdfDoc.Save(pdf_file_tosave, Aspose.Pdf.SaveFormat.Pdf);
                MessageBox.Show("Save Successfully!");
                pdfDoc.Dispose();
                Process.Start(pdf_file_tosave);
            }
            catch
            {
                MessageBox.Show("Faild to Save!");
            }
        }
        private void LoadImageOnPdfFile_3()
        {
            //create filename of pdf_file_tosave 
            var pdf_file_tosave = @"d:\1.pdf";

            //new pdf doc
            var pdfDoc = new Aspose.Pdf.Document();

            //editing
            pdfDoc.SetTitle("hello!");
            var pdfPage = pdfDoc.Pages.Add();
            
            //var te1 = new Aspose.Pdf.Text.TextSegment();
            pdfPage.AddImage(@"F:\MyDesktop\en-HandWritting\e.jpg", new Aspose.Pdf.Rectangle(0,0, 500, 500));

            //save
            try
            {
                pdfDoc.Save(pdf_file_tosave, Aspose.Pdf.SaveFormat.Pdf);
                MessageBox.Show("Save Successfully!");
                pdfDoc.Dispose();
                Process.Start(pdf_file_tosave);
            }
            catch
            {
                MessageBox.Show("Faild to Save!");
            }
        }
        private void CreateTextOnPdfFile_2()
        {
            //create filename of pdf_file_tosave 
            var pdf_file_tosave = @"d:\1.pdf";

            //new pdf doc
            var pdfDoc = new Aspose.Pdf.Document();

            //editing
            pdfDoc.SetTitle("hello!");
            var pdfPage = pdfDoc.Pages.Add();

            //var te1 = new Aspose.Pdf.Text.TextSegment();
            var te = new Aspose.Pdf.Text.TextFragment();
            te.Text = "1233";
            pdfPage.Paragraphs.Add(te);

            //save
            try
            {
                pdfDoc.Save(pdf_file_tosave, Aspose.Pdf.SaveFormat.Pdf);
                MessageBox.Show("Save Successfully!");
                pdfDoc.Dispose();
                Process.Start(pdf_file_tosave);
            }
            catch
            {
                MessageBox.Show("Faild to Save!");
            }
        }
        private void CreatePdfFile_1()
        {
            //create filename of pdf_file_tosave 
            var pdf_file_tosave = @"d:\1.pdf";

            //new pdf doc
            var pdfDoc = new Aspose.Pdf.Document();

            //editing
            pdfDoc.SetTitle("hello!");
            var pdfPage = pdfDoc.Pages.Add();

            //save
            try
            {
                pdfDoc.Save(pdf_file_tosave, Aspose.Pdf.SaveFormat.Pdf);
                MessageBox.Show("Save Successfully!");
                pdfDoc.Dispose();
                Process.Start(pdf_file_tosave);
            }
            catch
            {
                MessageBox.Show("Faild to Save!");
            }
        }
    }
}
