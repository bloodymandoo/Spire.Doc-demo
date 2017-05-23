using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
namespace WordToPDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Document doc = new Document();
            doc.LoadFromFile(@"Sample.doc");
            doc.SaveToFile("Sample.pdf", FileFormat.PDF);
            WordDocViewer("Sample.pdf");
        }
        private void WordDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }
    }
}
