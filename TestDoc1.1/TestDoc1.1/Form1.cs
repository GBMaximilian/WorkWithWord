using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;


namespace TestDoc1._1

{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            #region говно


            /*

            //var application = new Microsoft.Office.Interop.Word.Application();
            //Создаём новый Word.Application
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //Загружаем документ
            Microsoft.Office.Interop.Word.Document doc = null;

            //object fileName = "Здесь путь до файла Word формата *.doc";
            object fileName = @"D:\Корыто.docx";
            object falseValue = false;
            object trueValue = true;
            object missing = Type.Missing;

            doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);

            //Теперь у нас есть документ который мы будем менять.

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            //Задаём параметры замены и выполняем замену.
            object findText = "<Крыло>";
            object replaceWith = "<Н>";
            object replace = 2;

           // app.Selection.Find.Execute(ref findText, ref missing, ref missing, ref missing,
           // ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
            //ref replace, ref missing, ref missing, ref missing, ref missing);



            //Открываем документ для просмотра.
            app.Visible = true;
            */
            // Получить объект приложения Word.


            /*
                        Microsoft.Office.Interop.Word._Application word_app = new Microsoft.Office.Interop.Word.Application();

                        // Сделать Word видимым (необязательно).
                        word_app.Visible = true;

                        // Создаем документ Word.
                        object missing = Type.Missing;
                        Microsoft.Office.Interop.Word._Document word_doc = word_app.Documents.Add(
                            ref missing, ref missing, ref missing, ref missing);

                        // Создаем абзац заголовка.
                        Microsoft.Office.Interop.Word.Paragraph para = word_doc.Paragraphs.Add(ref missing);
                        para.Range.Text = "Кривая хризантемы";
                        object style_name = "Заголовок 1";
                        para.Range.set_Style(ref style_name);
                        para.Range.InsertParagraphAfter();

                        // Добавить текст.
                        para.Range.Text = "Сделать кривую хризантемы" +
                    "используйте следующие параметрические уравнения, когда t идет" +
                    "от 0 до 21 *? для генерации" +
                    "точки, а затем соединить их";
                        para.Range.InsertParagraphAfter();




                        object filename = Path.GetFullPath (
                    Path.Combine (Application.StartupPath, @"D:\")) +
                    "test.doc";
                word_doc.SaveAs(ref filename, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing);

                // Закрыть.
                object save_changes = false;
                word_doc.Close (ref save_changes, ref missing, ref missing);
                word_app.Quit (ref save_changes, ref missing, ref missing);


                         */
            #endregion

            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            string filename = openFileDialog1.FileName;


            var word_app = new Word.Application();
            word_app.Visible = false;//отображение ворда во время работы кода

            var wordDoc = word_app.Documents.Open(filename);

            SAVE(wordDoc, word_app);
        }
            

        private void button2_Click(object sender, EventArgs e)
        {
            var word_app = new Word.Application();

            word_app.Visible = false;

            // Создаем документ Word.
            object missing = Type.Missing;

            var word_doc = word_app.Documents.Add();

            // Создаем абзац заголовка.
            var para = word_doc.Paragraphs.Add(ref missing);

            object style_name = "Заголовок 1";
            para.Range.set_Style(ref style_name);
            para.Range.Text += "Кривая хризантемы";
            para.Range.InsertParagraphAfter();


            para.Range.Font.Size = 13;      
            para.Range.Font.Bold = -1;
            

            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            para.Range.Text += richTextBox1.Text;
            para.Range.InsertParagraphAfter();

            para.Range.Font.Italic = -1;
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            para.Range.Text += richTextBox1.Text;
            para.Range.InsertParagraphAfter();



            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            para.Range.Text += richTextBox1.Text;
            para.Range.InsertParagraphAfter();


            SAVE(word_doc, word_app);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Center;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Right;
        }


        private void SAVE(Word.Document word_doc, Word.Application word_app)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            string filename = saveFileDialog1.FileName;
            try
            {
                word_doc.SaveAs(filename);
                MessageBox.Show("файл сохранен");
                word_app.Visible = true;
            }
            catch { MessageBox.Show("произошла ошибка"); }
        }

        
    }
}
