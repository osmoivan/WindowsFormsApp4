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
using Word = Microsoft.Office.Interop.Word; //Введение алиаса пространства имен Word

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            textBox1.Text = "1111qwertyuiop[]asdfghjklzxcvbnm";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int[] x = new int[2];
            x = Ivan.SelectedTextIntoIndexForSemanticFragmentTable(textBox1);
            Ivan.AddIndexIntoSemanticFragmentTable("C://Users/ivan_/Desktop/time.txt");

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Sergey.ComboBox1_SelectedIndexChanged(comboBox1,textBox1);          //Переход между документами в комбобоксе
        }

        private void удалитьДокументToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Sergey.удалитьФайлToolStripMenuItem_Click(comboBox1, textBox1);    //Удаление текущего документа (по значению в комбобоксе)
        }

        private void добавитьДокументToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();             //Открытие диалогового окна для подгрузки текста из файла
            Sergey.добавитьФайлToolStripMenuItem_Click(textBox1, comboBox1, openFileDialog);
        }

        private void удалитьВсеДокументыToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Sergey.удалитьВсеДокументыToolStripMenuItem_Click(comboBox1, textBox1);  //Удаление всех существующих документов
        }
    }
}
public class Ivan
{
    private static int[] outdata = new int[2];
    public static int[] SelectedTextIntoIndexForSemanticFragmentTable(System.Windows.Forms.TextBox textBox)
    {
        if (textBox.SelectionLength > 0)
        {
            outdata[0] = textBox.SelectionStart;
            outdata[1] = textBox.SelectionStart + textBox.SelectionLength - 1;
            return outdata;
        }
        else
        {
            MessageBox.Show("Вы не выделили смысловой фрагмент");
            return outdata;
        }
    }
    private static int[,] ReadInformFromSemanticFragmentTable(string File)
    {
        int lengthTable = System.IO.File.ReadAllLines(File).Length + 1;
        int[,] index = new int[lengthTable, 2];
        int ind = 0;
        using (StreamReader sr = new StreamReader(File))   //System.IO.File.Create(File))
        {
            while (sr.Peek() >= 0)
            {
                string s = sr.ReadLine();
                string[] s1 = s.Split('\t');
                index[ind, 0] = Convert.ToInt32(s1[1]);
                index[ind, 1] = Convert.ToInt32(s1[2]);
                ind = ind + 1;
            }
        }
        return index;
    }
    private static void WriteInformIntoSemanticFragmentTable(string File, int[,] index)
    {
        using (StreamWriter sw = new StreamWriter(File))
        {
            for (int i = 0; i < index.Length / 2; i++)
            {
                sw.WriteLine("СФ" + i.ToString() + "\t" + index[i, 0].ToString() + "\t" + index[i, 1].ToString());
            }
        }
    }
    public static void AddIndexIntoSemanticFragmentTable(string File)
    {
        int[,] index = ReadInformFromSemanticFragmentTable(File);
        index = Ivan.SortMatrix(index);
        index[0, 0] = outdata[0];
        index[0, 1] = outdata[1];
        index=Ivan.CheckCrossingElements(index);
        WriteInformIntoSemanticFragmentTable(File, index);
    }
    private static int[,] SortMatrix(int[,] index)
    {
        List<int> FirstElement = new List<int>();
        List<int> SecondElement = new List<int>();
        int[,] timeindex = new int[index.GetLength(0), index.GetLength(1)];
        for (int i = 0; i < index.GetLength(0); i++)
        {
            FirstElement.Add(index[i, 0]);
            SecondElement.Add(index[i, 1]);
        }
        for (int i = 0; i < index.GetLength(0); i++)
        {
            int c = FirstElement.IndexOf(FirstElement.Min());
            timeindex[i, 0] = FirstElement[c];
            timeindex[i, 1] = SecondElement[c];
            FirstElement.RemoveAt(c);
            SecondElement.RemoveAt(c);
        }
        return timeindex;
    }
    //private static int[,] RemoveOneElement(int [,]index)
    //{

    //}
    private static int[,] CheckCrossingElements(int[,] index)
    {
        if (index.Length > 2)
        {
            //Если новый интервал полоностью покрывает старый
            if (index[0, 0]<=index[1,0] && index[1,0]>=index[0,0])
            {

            }
        }
        return index;
    }
}
public class Sergey
{
    static private Word.Application S_wordapp;
    static private Word.Document S_worddocument;
    static string S_namefiledirect;
    static string S_namefile;
    static string S_text;
    static List<string> S_nameFile = new List<string>();
    static List<string> S_textFile = new List<string>();

    public static void добавитьФайлToolStripMenuItem_Click(System.Windows.Forms.TextBox textBox, System.Windows.Forms.ComboBox comboBox, System.Windows.Forms.OpenFileDialog openFileDialog1)
    {
        
        if (openFileDialog1.ShowDialog() == DialogResult.OK)                           //Открытие диалогового окна для открытия файла
        {
            S_namefiledirect = openFileDialog1.FileName;                               //Сохранение в переменную пути к выбранному фалу
            S_namefile = openFileDialog1.SafeFileName;                                 //Сохранение в переменную имени выбранного файла
            S_wordapp = new Word.Application();                                        //Создаем объект Word - равносильно запуску Word.
            S_wordapp.Visible = false;
            Object filename = S_namefiledirect;
            S_worddocument = S_wordapp.Documents.Open(ref filename);                   //Открываем конкретный существующий word документ из нужной директории.
            Object begin = Type.Missing;                                               //В документе определяем диапазон,вызовом метода Range  
            Object end = Type.Missing;                                                 //с передачей ему начального 
            Word.Range wordrange = S_worddocument.Range(ref begin, ref end);           //и конечного значений позиций символов.
            wordrange.Copy();                                                          //Копирование в буфер обмена.
            S_text = Clipboard.GetText();                                              //Сохранение в переменную скопированного текста
            textBox.Text = S_text;                                                    //Извлекаем из буфера обмена копированный текст.
            comboBox.Items.Add(S_namefile);                                           //Добавление в combobox1 имени файла, с которого мы скопировали текст 
            comboBox.Text = S_namefile;                                               //Надпись комбобокса меняется на имя файла
            S_nameFile.Add(S_namefile);                                                //Добавление в список имени файла
            S_textFile.Add(S_text);                                                    //Добавление в список текста файла
            Object saveChanges = Word.WdSaveOptions.wdPromptToSaveChanges;
            Object originalFormat = Word.WdOriginalFormat.wdWordDocument;
            Object routeDocument = Type.Missing;
            S_wordapp.Quit(ref saveChanges, ref originalFormat, ref routeDocument);    // Закрытие файла
            S_wordapp = null;
        }
    }


    public static void ComboBox1_SelectedIndexChanged(System.Windows.Forms.ComboBox comboBox,System.Windows.Forms.TextBox textBox)
    {
        for (int i = 0; i < S_nameFile.Count; i++)                                 
        {
            if (comboBox.Text == S_nameFile[i])
            {
                textBox.Text = S_textFile[i];
            }
        }
    }

    public static void удалитьФайлToolStripMenuItem_Click(System.Windows.Forms.ComboBox comboBox,System.Windows.Forms.TextBox textBox)
    {
        for (int i = 0; i < S_nameFile.Count; i++)
        {
            if (MessageBox.Show("Удалить текущий документы?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                if (comboBox.Text == S_nameFile[i])
                {
                    comboBox.Items.Remove(S_nameFile[i]);
                    S_nameFile.RemoveAt(i);
                    S_textFile.RemoveAt(i);
                    comboBox.Text = "Файл";
                    textBox.Text = "";
            }
        }
    }
    public static void удалитьВсеДокументыToolStripMenuItem_Click(System.Windows.Forms.ComboBox comboBox,System.Windows.Forms.TextBox textBox)  
    {
        if (MessageBox.Show("Удалить все существующие документы?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
        {
            comboBox.Items.Clear();
            comboBox.Text = "Файл";
            S_nameFile.Clear();
            S_textFile.Clear();
            textBox.Text = "";
        }
    }
}
