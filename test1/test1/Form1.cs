using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Bookmark = Microsoft.Office.Interop.Word.Bookmark;

namespace test1
{
    public partial class Form1 : Form
    {
        private readonly string FileName = @"F:\РГУ\Учеба.VII-й семестр\Диплом\test1\test1\bin\Debug\test.docx";
        private Word.Application WordApp;
        private Word.Document WordDocument;
        private string[] bookmarks;
        private int[] bm_numb;
       // private string[] text = new string[] { "Дмитрий", " Силкин", "Владимирович", "21" };
        object missing = System.Reflection.Missing.Value;

        private Word.Paragraphs wordparagraphs;

        public Form1()
        {
            InitializeComponent();
        }

        public bool GoToBookMark(string bookMarkName)
        {

            if (WordDocument.Bookmarks.Exists(bookMarkName))
            {
                object what = Word.WdGoToItem.wdGoToBookmark;
                object name = bookMarkName;
                GoTo(what, missing, missing, name);
                return true;
            }
            return false;
        }

        public void GoTo(object what, object which, object count, object name)
        {
            WordApp.Selection.GoTo(ref what, ref which, ref count, ref name);
        }

        public void InsertText(string text)
        {
            WordApp.Selection.TypeText(text);
        }

        public void FillBookmarkByName()
        {
            string[] text = new string[dataGridView1.RowCount];
            for (int i = 0; i < bookmarks.Length; i++)
            {
                if (text[i] == null)
                    continue;
                else
                {
                    text[i] = "  " + dataGridView1.Rows[i].Cells[1].Value.ToString();
                    GoToBookMark(bookmarks[i]);
                    InsertText(text[i]);
                }
            }

        }

        public void InsertTextField() // хуй знает че за поля имеются ввиду здесь
        {
            string txt = "       !!!!!das";
            Object begin = bm_numb[1];
            Object end = bm_numb[1] + txt.Length;
            Word.Range rng = WordDocument.Range(begin, end);
            WordDocument.Fields.Add(rng,ref missing,txt);
           
        }

        public void InsertTextbox()
        {            
            Word.Shape textbox;
            textbox = WordDocument.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 20, 20, 100, 100);

            
        }


        public void SaveAs(string FileName)
        {
            if (WordDocument == null)
            {
                WordDocument = WordApp.ActiveDocument;
            }
            object objFileName = FileName;
            WordDocument.SaveAs(ref objFileName);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            WordApp = new Word.Application();
            WordApp.Visible = false;

            try
            {
                WordDocument = WordApp.Documents.Open(FileName);

                int ii = 0;
                var orderedBoomarks = WordDocument.Bookmarks.Cast<Bookmark>().OrderBy(d => d.Start).ToList();
                bm_numb = new int[orderedBoomarks.Count];

                foreach(Bookmark bookmark in orderedBoomarks)
                {
                    bm_numb[ii] =  bookmark.Range.Start;
                    dataGridView1.Rows.Add(bookmark.Name.ToString());
                    ii++;
                }
               
                
                // WordApp.Visible = true;
                
                bookmarks = new string[dataGridView1.RowCount];
                for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        bookmarks[i] = dataGridView1.Rows[i].Cells[0].Value.ToString();
                       // dataGridView1.Rows[i].Cells[1].Value = bookmarks[i];        
                    }
                //wordparagraphs = WordDocument.Paragraphs; определение количества параграфов
                //dataGridView1.Rows[0].Cells[1].Value = wordparagraphs.Count.ToString();

            }
           
            catch
            {
                MessageBox.Show("Фиаско");
                WordApp.Quit();
                WordApp = null;
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string filename = @"F:\РГУ\Учеба.VII-й семестр\Диплом\test1\test1\bin\Debug\result.docx";
            FillBookmarkByName();
            InsertTextbox();
            SaveAs(filename);
            
            if (WordDocument != null)
            {
                WordDocument.Close(ref missing, ref missing, ref missing);
                WordApp.Application.Quit(ref missing, ref missing, ref missing);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (WordDocument != null)
            {
                WordApp.Application.Quit(ref missing, ref missing, ref missing);
                dataGridView1.Rows.Clear();    
            }            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            
                WordDocument.Close(ref missing, ref missing, ref missing);
                WordApp.Quit(ref missing, ref missing, ref missing);
                WordApp = null;
                dataGridView1.Rows.Clear();
        }

       
    }
}
