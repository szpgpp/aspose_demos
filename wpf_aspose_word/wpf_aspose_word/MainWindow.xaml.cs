using System;
using System.Collections.Generic;
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

namespace wpf_aspose_word
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
            //CreatePdfFile_1();
            //CreateTextOnPdfFile_2();
            //CreateTableOnPdfFile_3();
            DocumentBuilder_4();
        }
        private void DocumentBuilder_4()
        {
            //create filename of pdf_file_tosave 
            var word_file_tosave = @"d:\1.docx";

            //new pdf doc
            //var word_doc = new Aspose.Words.Document();
            Aspose.Words.DocumentBuilder builder1 = new Aspose.Words.DocumentBuilder();
            builder1.InsertHtml("<H1>this is a testing!</H1>");
            var rhp = new Aspose.Words.Drawing.RelativeHorizontalPosition();
            var rvp = new Aspose.Words.Drawing.RelativeVerticalPosition();

            builder1.InsertShape(Aspose.Words.Drawing.ShapeType.StraightConnector1, rhp, 10, rvp, 10, 300, 300, Aspose.Words.Drawing.WrapType.None);
            builder1.InsertShape(Aspose.Words.Drawing.ShapeType.CurvedConnector3, rhp, 10, rvp, 10, 200, 300, Aspose.Words.Drawing.WrapType.None);
            builder1.InsertShape(Aspose.Words.Drawing.ShapeType.CurvedConnector4, rhp, 10, rvp, 10, 100, 300, Aspose.Words.Drawing.WrapType.None);
            builder1.InsertShape(Aspose.Words.Drawing.ShapeType.CurvedConnector5, rhp, 10, rvp, 10, 50, 300, Aspose.Words.Drawing.WrapType.None);

            //save
            try
            {
                //word_doc.Save(word_file_tosave, Aspose.Words.SaveFormat.Docx);
                builder1.Document.Save(word_file_tosave, Aspose.Words.SaveFormat.Docx);
                MessageBox.Show("Save Successfully!");
                //pdfDoc.Dispose();
                Process.Start(word_file_tosave);
            }
            catch
            {
                MessageBox.Show("Faild to Save!");
            }
        }
        private void CreateTableOnPdfFile_3()
        {
            //create filename of pdf_file_tosave 
            var word_file_tosave = @"d:\1.docx";

            //new pdf doc
            var word_doc = new Aspose.Words.Document();
            var table1 = new Aspose.Words.Tables.Table(word_doc);
            var row1 = new Aspose.Words.Tables.Row(word_doc);
            var cell1 = new Aspose.Words.Tables.Cell(word_doc);
            //var para1 = new Aspose.Words.Paragraph(word_doc);
            cell1.AppendChild(new Aspose.Words.Paragraph(word_doc));
            cell1.FirstParagraph.AppendChild(new Aspose.Words.Run(word_doc,"1223333"));
            row1.AppendChild(cell1);
            table1.AppendChild(row1);
            word_doc.FirstSection.Body.AppendChild(table1);

            //save
            try
            {
                word_doc.Save(word_file_tosave, Aspose.Words.SaveFormat.Docx);
                MessageBox.Show("Save Successfully!");
                //pdfDoc.Dispose();
                Process.Start(word_file_tosave);
            }
            catch
            {
                MessageBox.Show("Faild to Save!");
            }
        }

        private void CreateTextOnPdfFile_2()
        {
            //create filename of pdf_file_tosave 
            var word_file_tosave = @"d:\1.docx";

            //new pdf doc
            var word_doc = new Aspose.Words.Document();
            word_doc.FirstSection.Body.AppendParagraph("1234455");
            //MessageBox.Show(word_doc.Sections.Count.ToString());

            //save
            try
            {
                word_doc.Save(word_file_tosave, Aspose.Words.SaveFormat.Docx);
                MessageBox.Show("Save Successfully!");
                //pdfDoc.Dispose();
                Process.Start(word_file_tosave);
            }
            catch
            {
                MessageBox.Show("Faild to Save!");
            }
        }
        private void CreatePdfFile_1()
        {
            //create filename of pdf_file_tosave 
            var word_file_tosave = @"d:\1.docx";

            //new pdf doc
            var word_doc = new Aspose.Words.Document();

            //save
            try
            {
                word_doc.Save(word_file_tosave, Aspose.Words.SaveFormat.Docx);
                MessageBox.Show("Save Successfully!");
                //pdfDoc.Dispose();
                Process.Start(word_file_tosave);
            }
            catch
            {
                MessageBox.Show("Faild to Save!");
            }
        }
    }
}
