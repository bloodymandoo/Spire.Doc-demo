using System;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;

namespace CreateNew
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Create word document
            Document document = new Document();
            //Get a section
            Section section = document.AddSection();
            //Get a paragraph
            Paragraph para = section.AddParagraph();
            document.SaveToFile("Sample.doc", FileFormat.Doc);
            WordDocView("Sample.doc");
        }
        private void WordDocView(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Document document = new Document();
            document.LoadFromFile(@"Sample.doc");
            Paragraph para = document.Sections[0].Paragraphs[0];
            para.AppendText(textBox1.Text);
            document.SaveToFile("Sample.doc", FileFormat.Doc);
            WordDocView("Sample.doc");
        }
    }
}
