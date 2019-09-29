using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.IO;

using System.Windows.Forms;

namespace DocsGeneratorV2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            String nama, nrp, prakname;

            if(saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                String[] filenames = openFileDialog1.FileNames;
                String savePath = Path.GetDirectoryName(saveFileDialog1.FileName) + "/";

                var app = new Microsoft.Office.Interop.Word.Application();
                app.Visible = false;
                
                var doc = app.Documents.Add(Type.Missing);

                //set paper size
                doc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
                doc.PageSetup.BottomMargin = app.CentimetersToPoints((float)1);
                doc.PageSetup.LeftMargin = app.CentimetersToPoints((float)1);
                doc.PageSetup.TopMargin = app.CentimetersToPoints((float)1);
                doc.PageSetup.RightMargin = app.CentimetersToPoints((float)1);

                //set spacing after and before
                doc.Paragraphs.SpaceAfter = 0;
                doc.Paragraphs.SpaceBefore = 0;

                foreach (String s in openFileDialog1.FileNames)
                {
                    // Write Title
                    var para = doc.Paragraphs.Last;
                    var range = para.Range;
                    range.Font.Bold = 1;
                    range.Font.Underline = WdUnderline.wdUnderlineSingle;
                    range.Font.Size = 14;
                    range.Font.Name = "Courier New";
                    range.InsertAfter(Path.GetFileName(s));

                    //Write Enter
                    doc.Paragraphs.Add();
                    para = doc.Paragraphs.Last;
                    range = para.Range;
                    range.Font.Bold = 1;
                    range.Font.Size = 10;

                    //write content
                    doc.Paragraphs.Add();
                    para = doc.Paragraphs.Last;
                    range = para.Range;
                    range.Font.Bold = 0;
                    range.Font.Size = 10;
                    range.Font.Name = "Courier New";
                    range.Font.Underline = WdUnderline.wdUnderlineNone;
                    range.InsertAfter(File.ReadAllText(s));

                    doc.Paragraphs.Add();
                }

                // set Footer
                doc.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
                doc.ActiveWindow.ActivePane.Selection.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                Object currentPage = WdFieldType.wdFieldPage;
                doc.ActiveWindow.Selection.Fields.Add(doc.ActiveWindow.Selection.Range, ref currentPage);

                // set Header
                doc.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
                doc.ActiveWindow.ActivePane.Selection.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                Object currentPage2 = WdFieldType.wdFieldPage;
                doc.ActiveWindow.Selection.TypeText($"[{txtname.Text} - {txtnrp.Text} - {txtprak.Text}]");

                doc.SaveAs2(saveFileDialog1.FileName);
                doc.Close();
                app.Quit();

                MessageBox.Show("GENERATE SUCCESS!");
            }
        }
    }
}
