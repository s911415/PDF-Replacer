using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.xtra.iTextSharp.text.pdf.pdfcleanup;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDFReplacer
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

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            var path = openFileDialog1.FileName;
            textBox1.Text = path;
            PdfReader reader = new PdfReader(path);
            label3.Text = reader.NumberOfPages.ToString();
            var rows = dataGridView1.Rows;
            rows.Clear();

            for (int i = 1, j = reader.NumberOfPages; i <= j; i++)
            {
                rows.Add(i, "");
            }

            reader.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                return;

            var reader = new PdfReader(textBox1.Text);
            var dstPath = saveFileDialog1.FileName;
            var rows = dataGridView1.Rows;
            var bookmarks = SimpleBookmark.GetBookmark(reader);
            var dst = new PdfStamper(reader, new FileStream(dstPath, FileMode.Create))
            {
                XmpMetadata = reader.Metadata,
                Outlines = bookmarks
            };

            for (int i = 0, j = rows.Count; i < j; i++)
            {
                var cell = rows[i].Cells;
                var pageNo = (int)cell[0].Value;
                var imgPath = (string)cell[1].Value;

                if (imgPath != null && imgPath != "")
                {
                    var rect = reader.GetPageSize(pageNo);
                    Image image = Image.GetInstance(imgPath);
                    image.SetAbsolutePosition(0, 0);
                    image.ScaleAbsolute(rect);
                    var cleanUpLocations = new List<PdfCleanUpLocation> { new PdfCleanUpLocation(pageNo, rect, BaseColor.WHITE) };
                    new PdfCleanUpProcessor(cleanUpLocations, dst).CleanUp();
                    var template = PdfTemplate.CreateTemplate(dst.Writer, image.Width, image.Height);
                    template.AddImage(image);
                    dst.GetOverContent(pageNo).AddTemplate(template, 0, 0, true);
                }
            }

            dst.Close();
            reader.Close();
            MessageBox.Show("Done");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog2.ShowDialog();
        }

        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            var rows = dataGridView1.Rows;
            string[] tmp = openFileDialog2.FileNames;
            Array.Sort<string>(tmp);
            int i = 0;

            foreach (var str in openFileDialog2.FileNames)
            {
                if (i < rows.Count)
                {
                    rows[i].Cells[1].Value = str;
                }

                i++;
            }
        }
    }
}
